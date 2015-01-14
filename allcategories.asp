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

<link href="/<%=strInstallDir%>css/css.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
</head>
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
<%Dim ThisPage,xxsx,xsfs,pageid,TempText,weblink
xsfs=trim(request("xsfs"))
pageid=trim(request("pageid"))
If xsfs=""  Then xsfs=2
xxsx=trim(request("xxsx"))
if request("page")="" then
  ThisPage=1		
else
  ThisPage=request("page")
end If

If xsfs=1 or pageid<>"" then
dim thepage,currentpage,totalpage,pagename,page,ii,dcity,city_id
city_oneid=strint(request("city_oneid"))
city_twoid=strint(request("city_twoid"))
city_threeid=strint(request("city_threeid"))
if city_twoid=0 and city_threeid=0 then dcity="city_oneid="&city_oneid
if city_threeid=0 and city_twoid<>0 then dcity="city_oneid="&city_oneid&" and city_twoid="&city_twoid
if city_twoid<>0 and city_threeid<>0 then dcity="city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid
if city_oneid<>0 and city_twoid=0 then city_id="&city_oneid="&city_oneid
if city_oneid<>0 and city_twoid<>0 and city_threeid=0 then city_id="&city_oneid="&city_oneid&"&city_twoid="&city_twoid
if city_oneid<>0 and city_twoid<>0 and city_threeid<>0 then city_id="&city_oneid="&city_oneid&"&city_twoid="&city_twoid&"&city_threeid="&city_threeid
'=======交易类型
leixing=request("leixing")
lx=request("leixing")
If leixing="" Then
leixing=""
Else
leixing=" And leixing='"&leixing&"'"
End If
'========交易类型END
if city_oneid=0 then
sql="select * from DNJY_info where yz=1"&leixing&" and dqsj>"&DN_Now&" order by hfje/fbts "&DN_OrderType&",b "&DN_OrderType&",fbsj "&DN_OrderType&""
else
sql="select * from DNJY_info where "&dcity&" and yz=1"&leixing&" and dqsj>"&DN_Now&" order by hfje/fbts "&DN_OrderType&",b "&DN_OrderType&",fbsj "&DN_OrderType&""
end if             
set rs=server.createobject("adodb.recordset") 
rs.open sql,conn,1,1  
if rs.eof then                          
%>
<TABLE WIDTH=980 BORDER=0 CELLPADDING=0 CELLSPACING=0 align="center" bgcolor="ffffff">
	<TR>
	 <TD width="100%" height="200" style="BORDER-RIGHT: #91b9dd 1px dotted; BORDER-TOP: #91b9dd 1px dotted; BORDER-LEFT: #91b9dd 1px dotted; BORDER-BOTTOM: #91b9dd 1px dotted" colspan="4" align="center">
<FONT color='#FF0000'>暂无分类信息！</FONT>
	 </TD>
  </TR>
</TABLE>
<%
rs.close:set rs=nothing                            
else  
thepage=strint(request("pageid"))
currentpage = thepage
if currentpage <= 0 then currentpage=1
const maxperpage=20                        
rs.Pagesize = maxperpage                      
totalpage = rs.PageCount 
rs.absolutepage = currentpage 
n=rs.recordcount  
pagename="allcategories.asp"                   
page = PageSplit(currentpage,totalpage,pagename,city_id,lx)%>


<TABLE width="980" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="ffffff">
 <tr>
    <TD height="28" align="center" background="img/bg_tools.gif">
<%
If request("leixing")<>"" Then
lx=request("leixing")+"类"
Else
lx="全部类型"
End if
set rslx=server.createobject("adodb.recordset")
sqllx = "select leixing from DNJY_config "
rslx.open sqllx,conn,1,1
leixing=split(rslx("leixing"),"|")
response.write"按交易类型显示：<select name='leixing' onChange=""var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location=jmpURL;} else {this.selectedIndex=0 ;}"" ><option value='?leixing='''>选择交易类型</option>"
response.write"<option value='?leixing=&"&CT_ID&"&xsfs=1'>全部类型</option>"
for x=0 to ubound(leixing)
response.write"<option value='?leixing="&leixing(x)&"&"&CT_ID&"&xsfs=1'>"&leixing(x)&"</option>"
next
response.write"</select>"
rslx.close
set rslx=Nothing
response.write "&nbsp;&nbsp;&nbsp;<span class=""style1""><b>"&lx&"</b></span>&nbsp;&nbsp;&nbsp;"
%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;		
<A href="allcategories.asp?<%=L_id%>&xsfs=1"><FONT style="font-size:14px" color="#0000BF">全部信息</FONT></A>&nbsp;&nbsp;页次：<B><%=currentpage%>/<%=totalpage%></B>&nbsp;&nbsp;每页<B><%=maxperpage%></B>条&nbsp;&nbsp;共刊登<B><%=n%></B>条信息&nbsp;&nbsp;&nbsp;&nbsp;<%response.write page%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A href="allcategories.asp?xsfs=2&<%=L_ID%>"><font color="#0000ff"><b>简洁方式>>></b></font></a></TD>
  </TR>
</TABLE>
<TABLE  border="0" cellpadding="0" cellspacing="0" align="center" id="table23">
  <TR>
    <TD height="6"></TD>
  </TR>
</TABLE>
<TABLE WIDTH=980 BORDER=0 CELLPADDING=0 CELLSPACING=0 align="center" bgcolor="ffffff">
  <TR>

<%ii=0 
do while ii < maxperpage and not rs.eof%> 
<TD width="25%" valign="top">
      <TABLE width="220" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="BFBFBF">
        <TR> 
          <TD width="100%" height="81" align="center" bgcolor="#EFF7FE" background="img/ad_poto.gif"> 
             <TABLE width="100%" border="0" cellspacing="0" cellpadding="0" height="81">
              <TR>
                <TD height="18"> 
                  <TABLE width="100%" border="0" cellspacing="0" cellpadding="0">
                    <TR>
                      <TD width="29%" height="22">
                      </TD>
                      <TD width="71%" height="22"><SPAN style="font-size: 9pt;color:#FF0000">[ <%=rs("type_one")%>]</SPAN>
						<FONT color="#008000" style="font-size: 9pt"><%=mid(DateValue(rs("addsj")),6)%></FONT></TD>
                    </TR>
                  </TABLE>
                </TD>
              </TR>
              <TR>
                <TD height="31">
                  <TABLE width="100%" border="0" cellspacing="0" cellpadding="0">
                    <TR> 
                      <TD width="37%"></TD>
                      <TD width="63%"><FONT color="#FFFFFF">↑有效期<%=datediff("d",now(),rs("dqsj"))%>天↑</FONT> 
                      </TD>
                    </TR>
                  </TABLE>
                </TD>
              </TR>
              <TR>
                <TD>
                  <TABLE width="100%" border="0" cellspacing="0" cellpadding="0">
                    <TR> 
                      <TD height="25" valign="bottom" align="center"><%if rs("Readinfo")=1 Then Response.Write "<img src=""img/lock.jpg"" alt=""普通会员权限浏览"">"%><%if rs("Readinfo")=2 Then Response.Write "<img src=""img/lockvip.jpg"" alt=""VIP会员权限浏览"">"%>
                        <A title=<%=rs("biaoti")%> href="<%=x_path("",rs("id"),"",rs("Readinfo"))%>" target="_blank"> 
                          
                          <B><FONT style="font-size:14px" <%IF RS("a")<>"0" THEN%>COLOR="<%=rs("a")%>"<%end if%>><%=CutStr(rs("biaoti"),22,"...")%></FONT></B> 
                          
                        </A>
                      </TD>
                    </TR>
                  </TABLE>
                </TD>
              </TR>
            </TABLE>
          </TD> 
  </TR>                
 <TR>
     <TD width="100%" height="120" align="center" valign="middle" bgcolor="#FFFFFF" > 
            <%if rs("c")<>1 and rs("tupian")="0" then%>
            <TABLE border="0" width="100%" cellspacing="4" cellpadding="0">
              <TR>                                    
                <TD width="100%" style="word-break:break-all"> 
				<%Dim vip,memo,dianhua
memo=CutStr(Rs("memo"),150,"...")
dianhua=rs("dianhua")
If request.cookies("dick")("username")<>rs("username") Then'自己发布的全权阅读
vip=0'登录会员
If request.cookies("dick")("username")<>"" Then vip=conn.Execute("Select vip From DNJY_user Where username='"&request.cookies("dick")("username")&"'")(0)
If rs("Readinfo")=1 And request.cookies("dick")("username")="" Then
memo="<font color=#666666>会员权限阅读</font>"
dianhua="<font color=#666666>会员权限阅读</font>"
End If
If rs("Readinfo")=2 And vip<1 Then
memo="<font color=#666666>VIP会员权限阅读</font>"
dianhua="<font color=#666666>VIP会员权限阅读</font>"
End If
End If
response.write memo%><BR>
                  <FONT color="#009999">联系人:<%=rs("name")%><BR>
                  电话:<%=dianhua%></FONT></TD>  
              </TR>                                  
            </TABLE>                   
            <%else%>
            <A target="_blank" href="<%=TempText%>">
            <IMG src="<%if lcase(left(rs("tupian"),7))<>"http://" then%>UploadFiles/infopic/<%end if%><%=rs("tupian")%>" border="0" width="120" height="100"> 
			</A>
          <%end if%>          </TD>  
 </TR>
 <TR>                
          <TD width="100%" bgcolor="FCEFEA" align="center" height="22">
		  <FONT color="#FF0000">点击:<%=rs("llcs")%></FONT>&nbsp;&nbsp; 
            <%if rs("hfxx")=1 then%>
			<IMG src="img/shouy.gif" alt="竞价信息">
            <%else
			if rs("tj")=>0 then%>
            <IMG src="images/jian.gif" alt="推荐信息">
            <%else
			if rs("username")<>"" then%>
            <FONT color="#FF0000">会员发布</FONT> 
            <%else%>
            <FONT color="#998000">游客发布</FONT> 
          <%end if
			end if
			end if%>          </TD> 
 </TR>
 </TABLE>
      <TABLE  border="0" cellpadding="0" cellspacing="0" align="center" id="table23">
        <TR>
          <TD height="6"></TD>
        </TR>
      </TABLE>
    </TD>
<%ii=ii+1
if ii mod 4=0 then  
%></TR><TR><%       
end if
rs.movenext
loop  
if (n mod 4) <> 0 and not ii mod maxperpage = 0 Then 
for ii = 1 to (4 - (n mod 4))%>
<TD width="25%">
  <TABLE border="0" width="100%" cellpadding="1" cellspacing="1">
    <TR>
      <TD width="100%" height="24" align="center" bgColor="#ffe8e8" style="BORDER-RIGHT: #dd91a5 1px dotted; BORDER-TOP: #dd91a5 1px dotted; BORDER-LEFT: #dd91a5 1px dotted; BORDER-BOTTOM: #dd91a5 1px dotted"><B><FONT style="font-size:14px" ><FONT color="red">分类信息 针对性强</FONT></FONT></B></TD>
    </TR>
    <TR>
      <TD width="100%" height="127" style="BORDER-RIGHT: #dd91a5 1px dotted; BORDER-TOP: #dd91a5 1px dotted; BORDER-LEFT: #dd91a5 1px dotted; BORDER-BOTTOM: #dd91a5 1px dotted"><DIV align="center">
          <TABLE border="0" width="93%" cellspacing="0" cellpadding="0">
            <TR>
              <TD width="100%"><P style="margin-top: 2" align="center"><FONT color="red">茫茫商海何处有生意<BR>
                        <BR>
                分类信息处处是商机</FONT><BR>
                <BR>
                本栏目诚征各类供求信息</P></TD>
            </TR>
          </TABLE>
      </DIV></TD>
    </TR>
    <TR>
      <TD width="100%" style="BORDER-RIGHT: #92dcb9 1px dotted; BORDER-TOP: #92dcb9 1px dotted; BORDER-LEFT: #92dcb9 1px dotted; BORDER-BOTTOM: #92dcb9 1px dotted" align="center"><FONT color="#008000">http://<%=weblink%></FONT> <BR>
          <FONT color="#FF0000">本站信息</FONT> <FONT color="#008000">有效期还有365天</FONT> </TD>
    </TR>
  </TABLE></TD>
<%next                                                  
end if                                                                            
rs.close:set rs=nothing%> 

</TR> 
</TABLE>


 
<TABLE  border="0" cellpadding="0" cellspacing="0" align="center" id="table23">
  <TR>
    <TD height="6"></TD>
  </TR>
</TABLE>
<TABLE width="980" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="CCCCCC">
  <TR>
    <TD height="28" align="center" background="img/bg_tools.gif"> <A href="allcategories.asp?<%=L_id%>&xsfs=1"><FONT style="font-size:14px" color="#0000BF">全部信息</FONT></A>&nbsp;&nbsp;页次：<B><%=currentpage%>/<%=totalpage%></B>&nbsp;&nbsp;每页<B><%=maxperpage%></B>条&nbsp;&nbsp;共刊登<B><%=n%></B>条信息&nbsp;&nbsp;&nbsp;&nbsp;
    <%response.write page%></TD>
  </TR>
</TABLE>

<%end If
else%>
<TABLE border="0" width="980" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF" align=center>
    <tr><td class=td4 height="26" colspan="7" >
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
response.write"<option value='?leixing=&"&CT_ID&"'>全部类型</option>"
for x=1 to ubound(leixing)
response.write"<option value='?leixing="&leixing(x)&"&"&CT_ID&"'>"&leixing(x)&"</option>"
next
response.write"</select>"
rslx.close
set rslx=Nothing
response.write "&nbsp;&nbsp;&nbsp;<span class=""style1""><b>"&lx&"</b></span>&nbsp;&nbsp;&nbsp;"
%>			
	
	<span class=font1>&nbsp;<font face="Wingdings 2" size="4">R</font> 
图例：</span><font color="#008080">（<img border="0" src="images/num/pic.gif" width="13" height="13">-图片 <img border="0" src="images/num/jsq.gif" width="12" height="12">-置顶 <img border="0" src="images/jian.gif" width="15" height="15">-推荐 <img border="0" src="images/new.gif" >-最新）</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A href="allcategories.asp?xsfs=1&<%=L_ID%>"><font color="#0000ff"><b>详细方式>>></b></font></A></td></tr>                   
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
	<%=f_typeinfo(xxsx,c1,c2,c3,t1,t2,t3,strint(request("page")),1,0,30,102,102,50,400)%>
   
</TABLE>
<%end If%>
<!-- 这里是主题结束 --><!--#include file="footer.asp"-->
