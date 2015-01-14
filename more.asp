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
<%set rs=server.createobject("ADODB.recordset")
dim b,bb,tj,xxsx,tjxx,zxhf,rmxx,hfxxx,ThisPage,tw,s
tw=request("tw")
If tw="" Or tw=1 Then
tw=1
s=10
Else
s=30
End if
xxsx=trim(request("xxsx"))
tjxx=trim(request("tjxx"))
zxhf=trim(request("zxhf"))
rmxx=trim(request("rmxx"))
hfxxx=trim(request("hfxxx"))
if request("page")="" then
  ThisPage=1		
else
  ThisPage=request("page")
end if
%>
<link href="/<%=strInstallDir%>css/css.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {color:#FF0000}
-->
</style>
</head>

<body>
<table width="980"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="21"><img src="img/top_s3.gif" width="21" height="49"></td>
        <td width="90" background="img/top_s1.gif"><div align="center"><img src="img/fangda.gif" width="70" height="27"></div></td>
        <td background="img/top_s1.gif"><%=f_search(1,1)%> </td>
        <td width="5"><img src="img/top_s2.gif" width="5" height="49"></td>
      </tr>
    </table></td>
    <td width="5"> </td>
    <td width="220"><table width="100%"  border="0" cellspacing="0" cellpadding="0" style="table-layout:fixed;word-wrap:break-word;word-break:break-all">
      <tr>
        <td width="21"><img src="img/top_s4.gif" width="21" height="49"></td>
        <td width="194" height="49" background="img/top_s1.gif"> <%=f_news(c1,c2,c3,3,0,0,0,0,0,0,24,20,1,13,11,1)%></td>
        <td width="5"><img src="img/top_s2.gif" width="5" height="49"></td>
      </tr>
    </table></td>
  </tr>
</table>
      <table width="980" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0">

  <tr>
    <td width="215" valign="top"><!--#include file=left.asp--></td>
<td width="5">&nbsp;</td>
    <td width="760" valign="top" bgcolor="#FFFFFF"><table width="760" border="0" align="right" cellpadding="0"  bgcolor="#FFFFFF" class="ty1">

      <tr>
        <td valign="top" bgcolor="#FFFFFF"><table width="760" height="152" border="0" align="right" cellPadding="0" cellspacing="0" >		  
		

	<!---上部通栏广告开始---->   
  <table align="center" bgcolor="#ffffff">
  <tr>  
   <td ><script language=javascript src="<%=adjs_path("ads/js","fygg",c1&"_"&c2&"_"&c3)%>"></script></td>
  </tr>
  </table> 
  <!---上部通栏广告结束----> 

<tr>
   <td height="57" valign="top" width="760"><div align="right">
    <table width="100%" height="30" border="0" cellpadding="0" cellspacing="0" rules="all" id="DataGrid1">
    <tr><td class="ty2" height="26" colspan="7">
<%If request("leixing")<>"" Then
lx=request("leixing")+"类"
Else
lx="全部类型"
End if
set rslx=server.createobject("adodb.recordset")
sqllx = "select leixing from DNJY_config "
rslx.open sqllx,conn,1,1
leixing=split(rslx("leixing"),"|")
response.write"按交易类型显示：<select name='leixing' onChange=""var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location=jmpURL;} else {this.selectedIndex=0 ;}"" ><option value='?leixing='''>选择交易类型</option>"
response.write"<option value='?leixing=&"&CT_ID&"&xxsx="&xxsx&"&tw="&tw&"'>全部类型</option>"
for x=0 to ubound(leixing)
response.write"<option value='?leixing="&leixing(x)&"&"&CT_ID&"&xxsx="&xxsx&"&tw="&tw&"'>"&leixing(x)&"</option>"
next
response.write"</select>"
rslx.close
set rslx=Nothing
response.write "&nbsp;&nbsp;&nbsp;<span class=""style1""><b>"&lx&"</b></span>&nbsp;&nbsp;&nbsp;"
%>		
	
	
	<span class=font1>&nbsp;<font face="Wingdings 2" size="4">R</font> 
图例：</span><font color="#008080">（<img border="0" src="images/num/pic.gif" width="13" height="13">-图片 <img border="0" src="images/num/jsq.gif" width="12" height="12">-置顶 <img border="0" src="images/jian.gif" width="15" height="15">-推荐 <img border="0" src="images/new.gif" >-最新）</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%If tw=1 Then%><A href="?xxsx=<%=xxsx%>&tw=2&<%=L_ID%>"><font color="#0000ff"><b>简洁方式>>></b></font></A><%else%><A href="?xxsx=<%=xxsx%>&tw=1&<%=L_ID%>"><font color="#0000ff"><b>详细方式>>></b></font></A><%End if%></td></tr>                   
                  <tr align="middle" bgcolor="#FAFAFA" style="FONT-WEIGHT: bold; BORDER-LEFT-COLOR: #2E99CC; BORDER-BOTTOM-COLOR: #2E99CC; BORDER-TOP-COLOR: #2E99CC; BACKGROUND-COLOR: #2E99CC; BORDER-RIGHT-COLOR: #2E99CC" >
                   
                    <td height="26" bgcolor="#FAFAFA" style="width: 315; background-color:#FAFAFA"><span class="style1">信息标题</span></td>
                    <td height="26" style="width: 62; background-color:#FAFAFA"><p class="style1">价格</td>
                    <td height="26" style="width: 41; background-color:#FAFAFA"><span class="style1">点/回</span></td>
                    <td height="26" style="width: 93px; background-color:#FAFAFA"><span class="style1">发布日期</span></td>
                    <td style="width: 93px; background-color:#FAFAFA"><span class="style1">交易</span></td>
                    <td style="width: 93px; background-color:#FAFAFA"><span class="style1">有效期</span></td>
                  </tr>
                  <tr align="center" bgcolor="#FAFAFA">
                    <td width="33" height="1" bgcolor="#CCCCCC"></td>
                    <td width="315" height="1" bgcolor="#CCCCCC"></td>
                    <td width="62" height="1" bgcolor="#CCCCCC"></td>
                    <td width="41" height="1" bgcolor="#CCCCCC"></td>
                    <td width="93" height="1" colspan="3" bgcolor="#CCCCCC"></td>
                  </tr>
                
<%=f_typeinfo(xxsx,c1,c2,c3,t1,t2,t3,strint(request("page")),1,tw,s,102,102,30,400)%>
                  </table>
                </div></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
    <td width="4" bgcolor="#E6E6E6"></td>
  </tr>
</table>
</body>
</html>
<%Call CloseDB(conn)%>
<!--#include file=footer.asp-->