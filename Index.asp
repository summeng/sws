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
<link href="skin/ltby/css.css" type="text/css" rel="stylesheet">
<%dim b,bb,tj,dh
id=trim(request("id"))%>

<!--对联广告代码开始-->
<script src="<%=adjs_path("ads/js","dl",c1&"_"&c2&"_"&c3)%>"></script>
<!--对联广告代码结束-->
<!---全屏浮动广告开始---->
  <table align="center" bgcolor="#ffffff">
  <tr>
   <td ><script language=javascript src="<%=adjs_path("ads/js","top",c1&"_"&c2&"_"&c3)%>"></script></td>
  </tr>
  </table>
<!---全屏浮动广告结束---->

<body>
<script language="javascript">
//检验表单的合法性
function SearchInfoForm() {
	if (document.SearchInfo.Keyword.value == "") {
		alert("\关健字不能为空！");
		document.SearchInfo.Keyword.focus();
	}
	else {
        return true;
    }
    return false;
}
function resetSearchImg(){
	var fact = document.all.SearchInfo
	fact.Keyword.style.backgroundImage="";
}
</script>
<SCRIPT>
	function toueme(){
		document.getElementById("toubiao").style.display="none";
	}
</SCRIPT>
<SCRIPT>
	function toueme1(){
		document.getElementById("toubiao1").style.display="none";
	}
</SCRIPT>
<SCRIPT>
	function toueme2(){
		document.getElementById("toubiao2").style.display="none";
	}
</SCRIPT>
<SCRIPT>
	function toueme3(){
		document.getElementById("toubiao3").style.display="none";
	}
</SCRIPT>
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td  height="0" style=" border-style:none" align="center" width="980"><script language=javascript src="<%=adjs_path("ads/js","bigpic",c1&"_"&c2&"_"&c3)%>"></script></td>
  </tr>
</table>        

      <table width="980" border="0" align="center" cellspacing="0" cellpadding="0" class="tablest">
        <tr> 
          <td class="middle" width="980" height="45" align="center" valign="top"><div class="top"><span style="float:left;padding-top:9px;"><%=f_search(3,2)%></div></td><td width="20" height="22"></td><td width="208" height="52"><a href="<%if request.cookies("dick")("username")<>"" then %><%=adduserinfo%>?<%=CT_ID%><%Elseif addxinxi=1 then%><%=addinfo%>?<%=CT_ID%><%else%>level.asp?<%=CT_ID%><%End if%>"><img src="skin/ltby/addxx.jpg" width=208 height=52 border=0></a></td>
        </tr>
      </table>
<table width="10" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="5"></td>
  </tr>
</table>
<table width="980" height="222" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="200" valign="top">

     <table width="200"  height="125" cellpadding="1" cellspacing="1" class="tablest">
        <tr>
          <td height="30" class="tablest2" valign="top" >
            <div class="infoft10">网站公告</div>
          </td>
        </tr>
        <tr>
          <td class="tablest1"  valign="top"><%=announce(c1,c2,c3,1,0,1,0,8,1,24,13,20)%></td>
        </tr>
      </table>
      <table width="200" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table> 
      <table width="200"  cellpadding="1" cellspacing="1" class="tablest">
        <tr>
          <td height="30" class="tablest2" valign="top" ><A href="news.asp?n=1,0,0&<%=C_ID%>"><div class="infoft10">新闻动态</div></a></td>
        </tr>
        <tr>
          <td class="tablest1"  valign="top"><%=news(c1,c2,c3,0,1,0,0,1,0,0,0,0,0,10,1,24,13,20)%></td>
        </tr>
      </table>   
      <table width="200" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
      <table width="200"  cellpadding="1" cellspacing="1" class="tablest">
        <tr>
          <td height="30" class="tablest2" valign="top" ><A href="news.asp?n=2,0,0&<%=C_ID%>"><div class="infoft10">服务指南</div></a></td>
        </tr>
        <tr>
          <td class="tablest1"  valign="top"><%=news(c1,c2,c3,0,2,0,0,1,0,0,0,0,0,10,1,24,13,20)%></td>
        </tr>
      </table>  
      <table width="200" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
<%if lmkg13="1" then%>
      <table width="200"  cellpadding="1" cellspacing="1" class="tablest">
        <tr>
          <td height="30" class="tablest2" valign="top" ><a href="bmcslist.asp" >
            <div class="infoft10">店铺资源共享</div>
            </a></td>
        </tr>
        <tr>
          <td class="tablest1"  valign="top"><%=vipnews("",0,0,1,1,0,0,0,20,20,1,15,10,1)%></td>
        </tr>
      </table>
      <table width="200" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
<%end if%>

	  </td>
<!---左边栏结束----->
    <td width="7"></td>

<td width="773" valign="top">
<table width="773" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td width="500" height="192" valign="top">
<!---幻灯广告开始---->
<table width="100%" cellspacing="0" cellpadding="0" class="tablest">
<tr><td width="500" ><%if slidelx=0 then%>
<script src="<%=adjs_path("UploadFiles/slidea","slide",c1&"_"&c2&"_"&c3)%>"></script>
<%elseif slidelx=1 then%>
<script src="UploadFiles/js/hd.js"></script>
<%else%>
<%end if%>  </td>
</tr></table>
<table width="500" border="0" cellspacing="0" cellpadding="0">
<tr>
<td  height="12"></td>
</tr>
 </table>
<!---幻灯广告结束---->		  
<!---头条图片标题开始---->
<table width="100%" cellspacing="0" cellpadding="0" >
<tr><td width="500" > <%=f_tpbt(c1,c2,c3)%> </td>
</tr></table>
<!---头条图片标题结束---->
		  
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
            <table width="500" height="153"  align="center" cellpadding="0" cellspacing="0"  class="tablest">
              <tr>
                <td height="32" valign="top" style="border-style:none;" background="skin/ltby/titlebg1.png"><table width="515" height="32" border="0" align="center" cellpadding="0" cellspacing="0">
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
                <td height="270" style="border-style:none;" valign="top"><table id="mainTableq" width="500" border="0" align="center" cellpadding="0" cellspacing="0">
                    <!--最新信息-->
                    <TBODY style="DISPLAY: block">
                      <tr>
                        <td width="509" style="padding:6px;"><%=f_info(1,0,0,0,c1,c2,c3,1,1,1,1,1,1,1,1,1,13,40,8,1,13,11)%></td>
                      </tr>
                    </TBODY>
                    <!--竞价信-->
                    <TBODY style="DISPLAY: none">
                      <tr>
                        <td style="padding:6px;"><%=f_info(2,0,0,0,c1,c2,c3,1,1,1,1,1,1,1,1,1,13,38,8,1,13,11)%></td>
                      </tr>
                    </TBODY>
                    <!--推荐信息-->
                    <TBODY style="DISPLAY: none">
                      <tr>
                        <td style="padding:6px;" valign="top"><%=f_info(3,0,0,0,c1,c2,c3,1,1,1,1,1,1,1,1,1,13,40,8,1,13,11)%></td>
                      </tr>
                    </TBODY>
					<!--热门信息-->
                    <TBODY style="DISPLAY: none">
                      <tr>
                        <td style="padding:6px;" valign="top"><%=f_info(4,0,0,0,c1,c2,c3,1,1,1,1,1,1,1,1,1,13,40,8,1,13,11)%></td>
                      </tr>
                    </TBODY>
                  </table></td>
              </tr>
            </table></td>
          <td width="7"> </td>
          <td width="240" valign="top"><div align="right">

<table width="247" height="125" cellpadding="1" cellspacing="1" class="tablest">
        <tr>
          <td height="30" class="tablest2" valign="top" ><a href="#" >
            <div class="infoft10">会员登陆</div>
            </a></td>
        </tr>
        <tr>
          <td class="tablest1"  valign="top"><%call userlogin()%></td>
</tr>
</table>

      <table width="200" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
<table width="248" height="64"  cellpadding="1" cellspacing="1"  bgcolor="#FFFFFF" class="tablest">
<tr>
<td bgcolor="#EAF1E9" style="border-style:none;padding:3px;">
<%if hdlb=1 Then%>
<script src="<%=adjs_path("ads/js","video",c1&"_"&c2&"_"&c3)%>"></script>
<%elseIf hdlb=2 Then%>
<script src="/<%=strInstallDir%>UploadFiles/js/flv.js"></script>
<%else%>
<script src="<%=adjs_path("ads/js","240",c1&"_"&c2&"_"&c3)%>"></script> 
<%end if%>
</td></tr></table>

      <table width="200" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
<table width="230" height="64"  cellpadding="1" cellspacing="1"  bgcolor="#FFFFFF" class="tablest">
<tr>
<td bgcolor="#EAF1E9" style="border-style:none;padding:3px;"><%if lmkg11="1" then%>
<iframe marginwidth="0" marginheight="0" src="bd_hd_player.asp?MaxNum=10&PicWidth=238&PicHeight=200&Title=20&TextHeight=15&<%=CT_ID%>" frameborder="0" width="238" height="230" scrolling="No"  align="center"></iframe><%else%>
<script src="<%=adjs_path("ads/js","240",c1&"_"&c2&"_"&c3)%>"></script><%end if%></td>
</tr></table>
            </div></td>
        </tr>
      </table>

      <table width="200" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
      <table width="773" border="0" cellspacing="1" cellpadding="1" class="tablest">
        <tr>          
          <td class="tablest1" align="center"><script src="<%=adjs_path("ads/js","syzb730",c1&"_"&c2&"_"&c3)%>"></script></td>
        </tr>
      </table>
<%if lmkg7="1" then%>
 <!---店铺推荐开始---->
      <table width="200" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
      <table width="773" border="0" cellspacing="1" cellpadding="1" class="tablest">
        <tr>
          <td width="24" class="tablest3"> <a href="sdindex.asp?<%=CT_ID%>&ttop=3" target="_blank"><FONT face=黑体 color=#111177 size=4>优秀店铺展示</FONT></a></td>
          <td class="tablest1"><DIV id=demo style="OVERFLOW: hidden; WIDTH: 740; COLOR: #ffffff">
              <TABLE cellSpacing=0 cellPadding=0 align=left border=0 cellspace="0">
                <TBODY>
                  <TR>
                    <TD id=demo1 vAlign=top><%=f_dmtj(c1,c2,c3,25,10,20,11,9,1,0,1,140,120)%></TD>
                    <TD id=demo2 vAlign=top> </TD>
                  </TR>
                </TBODY>
              </TABLE>
            </DIV>
            <SCRIPT>
var speed3=25//速度数值越大速度越慢
demo2.innerHTML=demo1.innerHTML
function Marquee(){
if(demo2.offsetWidth-demo.scrollLeft<=0)
demo.scrollLeft-=demo1.offsetWidth
else{
demo.scrollLeft++
}
}
var MyMar=setInterval(Marquee,speed3)
demo.onmouseover=function() {clearInterval(MyMar)}
demo.onmouseout=function() {MyMar=setInterval(Marquee,speed3)}
</SCRIPT>
          </td>
        </tr>
      </table>
  <!---店铺推荐结束---->
<%end if%>

   </td>
  </tr>
</table>



<table width="980" height="222" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="200" valign="top">
 <!---目录列表开始---->
  <% if  lmkg4="1"  then %>
      <table width="200" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
      <table width="200"  cellpadding="1" cellspacing="1" class="tablest">
        <tr>
          <td height="30" class="tablest2" valign="top" ><A href="type_list.asp?<%=CT_ID%>"><div class="infoft10">分类栏目导航</div></a></td>
        </tr>
        <tr>
          <td class="tablest1"  valign="top"><%=s_more(0,1,1,0,20,1,0,1,11,9,10,8)%></td>
        </tr>
      </table>
<%end if%>
<!---目录列表结束---->
      <table width="200" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
<%if lmkg12="1" then%>
      <table width="200"  cellpadding="1" cellspacing="1" class="tablest">
        <tr>
          <td height="30" class="tablest2" valign="top" ><a href="bmcslist.asp" >
            <div class="infoft10">便民服务</div>
            </a></td>
        </tr>
        <tr>
          <td class="tablest1"  valign="top"><%=bmfw(c1,c2,c3,20,2,80,20,12,4,1,1,1,1)%></td>
        </tr>
      </table>
      <table width="200" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
<%end if%>

	  </td>
<!---左边栏结束----->
    <td width="7"></td>

<td width="773" valign="top">
<table width="773" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
      <table width="200" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
      <table width="773" border="0" cellspacing="0" cellpadding="0" class="tablest">
        <tr>
          <td  height="7" align="center"><%=xxfl(c1,c2,c3,1,12,10,2,38,1,1,1,0,5,0,0)%></td>
        </tr>
      </table>


    </td>
  </tr>
</table>



<%if lmkg5="1" then%>
<table width="10" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="3"></td>
  </tr>
</table>

<table width="980" border="1" align="center" cellpadding="2" cellspacing="2" class="tablest">
  <tr >
    <td width="980" height="28" style=" border-style:none" background="skin/ltby/centertitle4.png"><div  class="infoft10">竞价信息 <A href="list.asp?xxsx=2&<%=C_ID%>">更多>>></a></div></td>
  </tr>

  <tr>
    <td  height="2" style=" border-style:none" align="center" width="980"><%=fkss(c1,c2,c3,1,1,1,1,1,0,1,1,1,20,70,5,5,1)%></td>
  </tr>
</table>
<%end if%>
<%if lmkg14="1" then%>
<table width="10" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="3"></td>
  </tr>
</table>
<table width="980" border="1" align="center" cellpadding="2" cellspacing="2" class="tablest">
  <tr >
    <td width="980" height="28" style=" border-style:none" background="skin/ltby/centertitle4.png"><div  class="infoft10">新闻资讯相册 <A href="news_list.asp?<%=C_ID%>">更多>>></a></div></td>
  </tr>

  <tr>
    <td  height="2" style=" border-style:none" align="center" width="980"><%=news_photo(c1,c2,c3,0,0,0,1,1,1,1,20,150,140,6,6,0,request("page"),"","")%></td>
  </tr>
</table>
<%end if%>
<%if lmkg15="1" then%>
<table width="10" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="3"></td>
  </tr>
</table>
<table width="980" border="1" align="center" cellpadding="2" cellspacing="2" class="tablest">
  <tr >
    <td width="980" height="28" style=" border-style:none" background="skin/ltby/centertitle4.png"><div  class="infoft10">电子杂志 <A href="magazine/magazine_show.asp?<%=C_ID%>">更多>>></a></div></td>
  </tr>

  <tr>
    <td  height="2" style=" border-style:none" align="center" width="980"><%=magazine(c1,c2,c3,6,18,180,180,20)%></td>
  </tr>
</table>
<%end if%>
<%if lmkg6="1" then%>
<table width="10" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="3"></td>
  </tr>
</table>

<table width="980" border="1" align="center" cellpadding="2" cellspacing="2" class="tablest">
  <tr >
    <td width="980" height="28" style=" border-style:none" background="skin/ltby/centertitle4.png"><div  class="infoft10">企业名片 <A href="qyindex.asp?<%=CT_ID%>&ttop=4">更多>>></a></div></td>
  </tr>

  <tr>
    <td  height="2" style=" border-style:none" align="center" width="980"><%=qymp(c1,c2,c3,6,9,11,20)%></td>
  </tr>
</table>
<%end if%> 
<table width="10" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="3"></td>
  </tr>
</table>
<table width="980" border="0" align="center" cellpadding="2" cellspacing="2">
  <tr>
    <td  height="0" style=" border-style:none" align="center" width="980"><script language=javascript src="<%=adjs_path("ads/js","bottom",c1&"_"&c2&"_"&c3)%>"></script></td>
  </tr>
</table>
<% if  lmkg8="1"  then %>
<table width="10" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="3"></td>
  </tr>
</table>
<table width="980" border="1" align="center" cellpadding="2" cellspacing="2" class="tablest">
  <tr >
    <td width="980" height="28" style=" border-style:none" background="skin/ltby/centertitle4.png"><div  class="infoft10">都市114 <a href="help.asp?id=4&<%=C_ID%>" target=_blank>我要订座</a>  <A href="114list.asp?<%=CT_ID%>">更多>>></a></div></td>
  </tr>
  <tr>
    <td  height="2" style=" border-style:none" width="980" align="center"><%=ds114(c1,c2,c3,5,35)%></td>
  </tr>
</table>
<%end if%>

<table width="10" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="3"></td>
  </tr>
</table>
<% if  lmkg9="1"  then %>
<table width="980" border="1" align="center" cellpadding="2" cellspacing="2" class="tablest">
  <tr>
    <td  height="2" style=" border-style:none"  width="980"><script src="UploadFiles/adfp/sinaflashfp.js"></script></td>
  </tr>
</table>
<table width="10" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="6"></td>
  </tr>
</table>
<%end if%>
<!----友情链接开始----->
<% if  lmkg10="1"  then %>
<table width="980" border="1" align="center" cellpadding="2" cellspacing="2" class="tablest">
  <tr >
    <td width="980" height="28" style=" border-style:none" background="skin/ltby/centertitle4.png"><div  class="infoft10">友情链接 <a href="user/link_edit.asp?<%=C_ID%>" target="_blank" >修改</a>            <A href="user/link_gd.asp?<%=C_ID%>" target=_blank>更多>>></a></div></td>
  </tr>
  <tr>
    <td width="980"  height="2" align="left" bgcolor="#F4F9FD" style=" border-style:none"><%=f_link(c1,c2,c3,0,8,88,10,50,2,20,50)%></td>
  </tr>
</table>
<%end if%>
<!---友情链接结束---->
<%If kefu=2 Or  kefu=4 then%>
<script>document.write("<scr"+"ipt language=\"javascript\" src=\"http://www.53kf.com/kf.php?arg=<%=kefuid%>&style=1&keyword="+escape(document.referrer)+"\"></scr"+"ipt>");</script>
<!-- 为了不影响您网站的加载速度，请将代码插到 </body> 的前一行 -->
<%elseif kefu=3 then%>
<script language="javascript" src="UploadFiles/js/54kefu.js" charset="utf-8"></script>
<%End if%> 
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="980"  align="center"><!--广告代码开始-->
<script src="<%=adjs_path("ads/js","syzb",c1&"_"&c2&"_"&c3)%>"></script>
<!--广告代码结束--></td>
  </tr>
</table>

<!--#include file="footer.asp" -->
<!--#include file="qq_kefu.asp"-->
<script language='javascript' src='/Resetting.js'></script>