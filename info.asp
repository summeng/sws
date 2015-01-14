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
<%set rs=server.createobject("ADODB.recordset")
dim b,bb,tj,xxsx,dick,ThisPage
xxsx=trim(request("xxsx"))
If xxsx="" Then xxsx=0
if request("page")="" then
ThisPage=1		
else
ThisPage=request("page")
end if%>
<link href="skin/ltby/css.css" type="text/css" rel="stylesheet">
<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
<body>
<table width="980"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="21"><img src="img/top_s3.gif" width="21" height="49"></td>
        <td width="90" background="img/top_s1.gif"><div align="center"><img src="img/fangda.gif" width="70" height="27"></div></td>
        <td background="img/top_s1.gif"><%=f_search(1,1)%></td>
        <td width="5"><img src="img/top_s2.gif" width="5" height="49"></td>
      </tr>
    </table></td>
    <td width="5"> </td>
    <td width="220"><table width="100%"  border="0" cellspacing="0" cellpadding="0" style="table-layout:fixed;word-wrap:break-word;word-break:break-all">
      <tr>
        <td width="21"> <a href="<%if request.cookies("dick")("username")<>"" then %><%=adduserinfo%>?<%=CT_ID%><%Elseif addxinxi=1 then%><%=addinfo%>?<%=CT_ID%><%else%>Type.asp?<%=CT_ID%><%End if%>"><img src="skin/ltby/addxx.jpg" width=190 height=52 border=0></a> </td>
      </tr>
    </table></td>
  </tr>
</table>
<table style="margin-bottom:8px;" width="980" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="CCCCCC">
  <tr>
    <td height="25" bgcolor="FFFAE6" style="padding-left:6px;" class="td2"> 当前位置：网站首页 ><strong><%=belongtype%></strong>> 信息列表</td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF" style="padding:8px;"><%=F_vassal_c_t(c1,c2,c3,t1,t2,t3,0,1,0,8,22,1)%></td>
  </tr>
</table>
<table style="margin-bottom:8px;" width="980" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="CCCCCC">
  <tr>
    <td height="25" bgcolor="FFFAE6" style="padding-left:6px;" class="td2"> 地区导航：<a href="list.asp?c=0,0,0&<%=T_ID%>">总站-></a></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF" style="padding:8px;"><%=F_vassal_c_t(c1,c2,c3,t1,t2,t3,0,1,0,10,22,0)%></td>
  </tr>
</table>
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0">

  <tr>
    <td width="200" height="100%" valign="top">

     <table width="200"  height="125" cellpadding="1" cellspacing="1" class="tablest">
        <tr>
          <td height="30" class="tablest2" valign="top" >
            <div class="infoft10">竞价信息</div>
          </td>
        </tr>
        <tr>
          <td class="tablest1"  valign="top"><%=f_info(2,0,0,0,c1,c2,c3,0,0,0,0,0,1,0,1,0,0,20,14,1,13,11)%></td>
        </tr>
      </table>
      <table width="200" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
      <table width="200" border="1" cellpadding="0" cellspacing="0" class="tablest" align="center">
        <tr>
          <td  style="border-style:none; padding:1px;" valign="top"><script language=javascript src="<%=adjs_path("ads/js","zuo1",c1&"_"&c2&"_"&c3)%>"></script></td>
        </tr>
      </table>
      <table width="200" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
     <table width="200"  height="125" cellpadding="1" cellspacing="1" class="tablest">
        <tr>
          <td height="30" class="tablest2" valign="top" >
            <div class="infoft10">最新信息</div>
          </td>
        </tr>
        <tr>
          <td class="tablest1"  valign="top"><%=f_info(1,0,0,0,c1,c2,c3,0,0,1,0,0,0,0,1,0,0,20,14,1,13,11)%></td>
        </tr>
      </table>
      <table width="200" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
      <table width="200" border="1" cellpadding="0" cellspacing="0" class="tablest" align="center">
        <tr>
          <td  style="border-style:none; padding:1px;" valign="top"><script language=javascript src="<%=adjs_path("ads/js","zuo2",c1&"_"&c2&"_"&c3)%>"></script></td>
        </tr>
      </table>
      <table width="200" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
     <table width="200"  height="125" cellpadding="1" cellspacing="1" class="tablest">
        <tr>
          <td height="30" class="tablest2" valign="top" >
            <div class="infoft10">推荐信息</div>
          </td>
        </tr>
        <tr>
          <td class="tablest1"  valign="top"><%=f_info(3,0,0,0,c1,c2,c3,0,1,0,0,0,0,0,1,0,0,20,14,1,13,11)%></td>
        </tr>
      </table>
      <table width="200" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
      <table width="200" border="1" cellpadding="0" cellspacing="0" class="tablest" align="center">
        <tr>
          <td  style="border-style:none; padding:1px;" valign="top"><script language=javascript src="<%=adjs_path("ads/js","zuo3",c1&"_"&c2&"_"&c3)%>"></script></td>
        </tr>
      </table>
      <table width="200" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
     <table width="200"  height="125" cellpadding="1" cellspacing="1" class="tablest">
        <tr>
          <td height="30" class="tablest2" valign="top" >
            <div class="infoft10">热门信息</div>
          </td>
        </tr>
        <tr>
          <td class="tablest1"  valign="top"><%=f_info(4,0,0,0,c1,c2,c3,0,0,0,0,0,0,0,1,0,1,20,14,1,13,11)%></td>
        </tr>
      </table>

</td> 
<td width="10"></td>
<td width="770" valign="top" bgcolor="#FFFFFF">
<!---上部通栏广告开始----> 
  <table align="center" bgcolor="#ffffff">
  <tr>  
   <td ><script language=javascript src="<%=adjs_path("ads/js","fygg",c1&"_"&c2&"_"&c3)%>"></script></td>
  </tr>
  </table>
  <!---上部通栏广告结束----> 
<table width="770"  border="0" align="right" cellPadding="0" cellspacing="0"  style="border-collapse: collapse"> 

          <tr>
            <td valign="top" width="770">
<div align="right">
                <table width="100%" height="30" border="0" cellpadding="0" cellspacing="0" rules="all" id="DataGrid1" class="tablest">
 <tr><td class="tablest2"  colspan="5" >
<%If request("leixing")<>"" Then
lx=request("leixing")+"类"
Else
lx="全部类型"
End if
Call OpenConn
set rslx=server.createobject("adodb.recordset")
sqllx = "select leixing from DNJY_config "
rslx.open sqllx,conn,1,1
leixing=split(rslx("leixing"),"|")
response.write"按交易类型显示：<select name='leixing' onChange=""var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location=jmpURL;} else {this.selectedIndex=0 ;}"" ><option value='?leixing='''>选择交易类型</option>"
response.write"<option value='?leixing=&"&CT_ID&"'>全部类型</option>"
for x=0 to ubound(leixing)
response.write"<option value='?leixing="&leixing(x)&"&"&CT_ID&"'>"&leixing(x)&"</option>"
next
response.write"</select>"
rslx.close
set rslx=Nothing
response.write "&nbsp;&nbsp;&nbsp;<span class=""style1""><b>"&lx&"</b></span>&nbsp;&nbsp;&nbsp;"
%>	
<span class=font1> <font face="Wingdings 2" size="4">R</font> 
图例</span><font color="#008080">(<img border="0" src="images/num/pic.gif" width="13" height="13">-图片 <img border="0" src="images/num/jsq.gif" width="12" height="12">-置顶 <img border="0" src="images/jian.gif" width="15" height="15">-推荐 <img border="0" src="images/new.gif" >-最新)</font></td><td class="tablest2" height="26" colspan="1" align="right"><a href="rss.asp?<%=CT_ID%>" target="_blank"><img border="0" src="img/rss.gif" alt="订阅本分类的最新信息" width="31" height="15"></a></td></tr>                   
                  <tr align="middle" bgcolor="#FAFAFA" style="FONT-WEIGHT: bold; BORDER-LEFT-COLOR: #2E99CC; BORDER-BOTTOM-COLOR: #2E99CC; BORDER-TOP-COLOR: #2E99CC; BACKGROUND-COLOR: #2E99CC; BORDER-RIGHT-COLOR: #2E99CC">                    
                    <td height="26" bgcolor="#FAFAFA" style="width: 350; background-color:#FAFAFA"><span class="style1">信息标题</span></td>
                    <td height="26" style="width: 62; background-color:#FAFAFA"><p class="style1">价格</td>
                    <td height="26" style="width: 41; background-color:#FAFAFA"><span class="style1">点/回</span></td>
                    <td height="26" style="width: 93px; background-color:#FAFAFA"><span class="style1">发布日期</span></td>
                    <td style="width: 93px; background-color:#FAFAFA"><span class="style1">交易</span></td>
                    <td style="width: 93px; background-color:#FAFAFA"><span class="style1">有效期</span></td>
                  </tr>
           <tr align="center" bgcolor="#FAFAFA">
           <td width="100%" height="1" bgcolor="#CCCCCC" colspan="6"></td>              
	</tr>    			
			
			<%=f_typeinfo(xxsx,c1,c2,c3,t1,t2,t3,strint(request("page")),1,1,10,102,102,30,400)%>
			</table>
                </div></td>
          </tr>
        </table></td>
  </tr>
</table><%Call CloseDB(conn)%>
<!--#include file="footer.asp" -->