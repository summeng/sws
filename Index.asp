<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file="top.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������Ұ�Ȩ����
'�ٷ��ͷ����� http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<!-- [DNJY] -->
<link href="skin/ltby/css.css" type="text/css" rel="stylesheet">
<%dim b,bb,tj,dh
id=trim(request("id"))%>

<!--���������뿪ʼ-->
<script src="<%=adjs_path("ads/js","dl",c1&"_"&c2&"_"&c3)%>"></script>
<!--�������������-->
<!---ȫ��������濪ʼ---->
  <table align="center" bgcolor="#ffffff">
  <tr>
   <td ><script language=javascript src="<%=adjs_path("ads/js","top",c1&"_"&c2&"_"&c3)%>"></script></td>
  </tr>
  </table>
<!---ȫ������������---->

<body>
<script language="javascript">
//������ĺϷ���
function SearchInfoForm() {
	if (document.SearchInfo.Keyword.value == "") {
		alert("\�ؽ��ֲ���Ϊ�գ�");
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
            <div class="infoft10">��վ����</div>
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
          <td height="30" class="tablest2" valign="top" ><A href="news.asp?n=1,0,0&<%=C_ID%>"><div class="infoft10">���Ŷ�̬</div></a></td>
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
          <td height="30" class="tablest2" valign="top" ><A href="news.asp?n=2,0,0&<%=C_ID%>"><div class="infoft10">����ָ��</div></a></td>
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
            <div class="infoft10">������Դ����</div>
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
<!---���������----->
    <td width="7"></td>

<td width="773" valign="top">
<table width="773" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td width="500" height="192" valign="top">
<!---�õƹ�濪ʼ---->
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
<!---�õƹ�����---->		  
<!---ͷ��ͼƬ���⿪ʼ---->
<table width="100%" cellspacing="0" cellpadding="0" >
<tr><td width="500" > <%=f_tpbt(c1,c2,c3)%> </td>
</tr></table>
<!---ͷ��ͼƬ�������---->
		  
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
                      <td width="105" height="30" class="sec22" style="border-style:none" onmouseover=secBoardq(0) ><a href="list.asp?<%=CT_ID%>&xxsx=1"><div align="center">������Ϣ </div></a></td>
                      <td width="105" class="sec11" style="border-style:none" onmouseover=secBoardq(1)><a href="list.asp?<%=CT_ID%>&xxsx=2"><div align="center">������Ϣ</div></a></td>
                      <td width="105" class="sec11" style="border-style:none" onmouseover=secBoardq(2)><a href="list.asp?<%=CT_ID%>&xxsx=3"><div align="center">�Ƽ���Ϣ</div></a></td>
                    <td width="105" class="sec11" style="border-style:none" onmouseover=secBoardq(3)><a href="list.asp?<%=CT_ID%>&xxsx=4"><div align="center">������Ϣ</div></a></td>
                      <td width="200" class="sec11" style="border-style:none"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr>
                <td height="270" style="border-style:none;" valign="top"><table id="mainTableq" width="500" border="0" align="center" cellpadding="0" cellspacing="0">
                    <!--������Ϣ-->
                    <TBODY style="DISPLAY: block">
                      <tr>
                        <td width="509" style="padding:6px;"><%=f_info(1,0,0,0,c1,c2,c3,1,1,1,1,1,1,1,1,1,13,40,8,1,13,11)%></td>
                      </tr>
                    </TBODY>
                    <!--������-->
                    <TBODY style="DISPLAY: none">
                      <tr>
                        <td style="padding:6px;"><%=f_info(2,0,0,0,c1,c2,c3,1,1,1,1,1,1,1,1,1,13,38,8,1,13,11)%></td>
                      </tr>
                    </TBODY>
                    <!--�Ƽ���Ϣ-->
                    <TBODY style="DISPLAY: none">
                      <tr>
                        <td style="padding:6px;" valign="top"><%=f_info(3,0,0,0,c1,c2,c3,1,1,1,1,1,1,1,1,1,13,40,8,1,13,11)%></td>
                      </tr>
                    </TBODY>
					<!--������Ϣ-->
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
            <div class="infoft10">��Ա��½</div>
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
 <!---�����Ƽ���ʼ---->
      <table width="200" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
      <table width="773" border="0" cellspacing="1" cellpadding="1" class="tablest">
        <tr>
          <td width="24" class="tablest3"> <a href="sdindex.asp?<%=CT_ID%>&ttop=3" target="_blank"><FONT face=���� color=#111177 size=4>�������չʾ</FONT></a></td>
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
var speed3=25//�ٶ���ֵԽ���ٶ�Խ��
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
  <!---�����Ƽ�����---->
<%end if%>

   </td>
  </tr>
</table>



<table width="980" height="222" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="200" valign="top">
 <!---Ŀ¼�б�ʼ---->
  <% if  lmkg4="1"  then %>
      <table width="200" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
      <table width="200"  cellpadding="1" cellspacing="1" class="tablest">
        <tr>
          <td height="30" class="tablest2" valign="top" ><A href="type_list.asp?<%=CT_ID%>"><div class="infoft10">������Ŀ����</div></a></td>
        </tr>
        <tr>
          <td class="tablest1"  valign="top"><%=s_more(0,1,1,0,20,1,0,1,11,9,10,8)%></td>
        </tr>
      </table>
<%end if%>
<!---Ŀ¼�б����---->
      <table width="200" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
<%if lmkg12="1" then%>
      <table width="200"  cellpadding="1" cellspacing="1" class="tablest">
        <tr>
          <td height="30" class="tablest2" valign="top" ><a href="bmcslist.asp" >
            <div class="infoft10">�������</div>
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
<!---���������----->
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
    <td width="980" height="28" style=" border-style:none" background="skin/ltby/centertitle4.png"><div  class="infoft10">������Ϣ <A href="list.asp?xxsx=2&<%=C_ID%>">����>>></a></div></td>
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
    <td width="980" height="28" style=" border-style:none" background="skin/ltby/centertitle4.png"><div  class="infoft10">������Ѷ��� <A href="news_list.asp?<%=C_ID%>">����>>></a></div></td>
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
    <td width="980" height="28" style=" border-style:none" background="skin/ltby/centertitle4.png"><div  class="infoft10">������־ <A href="magazine/magazine_show.asp?<%=C_ID%>">����>>></a></div></td>
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
    <td width="980" height="28" style=" border-style:none" background="skin/ltby/centertitle4.png"><div  class="infoft10">��ҵ��Ƭ <A href="qyindex.asp?<%=CT_ID%>&ttop=4">����>>></a></div></td>
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
    <td width="980" height="28" style=" border-style:none" background="skin/ltby/centertitle4.png"><div  class="infoft10">����114 <a href="help.asp?id=4&<%=C_ID%>" target=_blank>��Ҫ����</a>  <A href="114list.asp?<%=CT_ID%>">����>>></a></div></td>
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
<!----�������ӿ�ʼ----->
<% if  lmkg10="1"  then %>
<table width="980" border="1" align="center" cellpadding="2" cellspacing="2" class="tablest">
  <tr >
    <td width="980" height="28" style=" border-style:none" background="skin/ltby/centertitle4.png"><div  class="infoft10">�������� <a href="user/link_edit.asp?<%=C_ID%>" target="_blank" >�޸�</a>            <A href="user/link_gd.asp?<%=C_ID%>" target=_blank>����>>></a></div></td>
  </tr>
  <tr>
    <td width="980"  height="2" align="left" bgcolor="#F4F9FD" style=" border-style:none"><%=f_link(c1,c2,c3,0,8,88,10,50,2,20,50)%></td>
  </tr>
</table>
<%end if%>
<!---�������ӽ���---->
<%If kefu=2 Or  kefu=4 then%>
<script>document.write("<scr"+"ipt language=\"javascript\" src=\"http://www.53kf.com/kf.php?arg=<%=kefuid%>&style=1&keyword="+escape(document.referrer)+"\"></scr"+"ipt>");</script>
<!-- Ϊ�˲�Ӱ������վ�ļ����ٶȣ��뽫����嵽 </body> ��ǰһ�� -->
<%elseif kefu=3 then%>
<script language="javascript" src="UploadFiles/js/54kefu.js" charset="utf-8"></script>
<%End if%> 
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="980"  align="center"><!--�����뿪ʼ-->
<script src="<%=adjs_path("ads/js","syzb",c1&"_"&c2&"_"&c3)%>"></script>
<!--���������--></td>
  </tr>
</table>

<!--#include file="footer.asp" -->
<!--#include file="qq_kefu.asp"-->
<script language='javascript' src='/Resetting.js'></script>