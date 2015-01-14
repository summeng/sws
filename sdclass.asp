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
<style>
<!--
.LoginInput {
	BORDER-RIGHT: #eeeeee 1px solid; BORDER-TOP: #999999 1px solid; FONT: 12px "宋体", "Verdana", "Arial"; BORDER-LEFT: #999999 1px solid; BORDER-BOTTOM: #eeeeee 1px solid; HEIGHT: 17px; BACKGROUND-COLOR: #f6f6f6
}
-->
</style>
</head>

<body topmargin="0">
<div align="center">
<table style="margin-bottom:8px;" width="980" border="0" align="center" cellpadding="0" cellspacing="1" class="dq1">
  <tr>
    <td height="25" class="dq2" style="padding-left:6px;"> 地区导航：<a href="sdclass.asp?c=0,0,0&<%=T_ID%>">总站-></a></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF" style="padding:8px;"><%=W_vassal_c_t(0,1,c1,c2,c3,t1,t2,t3,0,1,0,10,0,22)%></td>
  </tr>
</table>
</div>

<table width="980" border="0" align="center" cellpadding="0" cellspacing="1" class="dq1">
  <tr>
    <td ><script src="<%=adjs_path("ads/js","toppic",c1&"_"&c2&"_"&c3)%>"></script></td>
  </tr>
</table>
<table  border="0" cellspacing="0" cellpadding="0">
    <tr>
     <td  height="7"></td>
    </tr>
</table>
<table style="margin-bottom:8px;" width="980" border="0" align="center" cellpadding="0" cellspacing="1" class="dq1">
  <tr>
    <td height="25" class="dq4" style="padding-left:6px;"><b>店 铺 分 类</b> &nbsp;&nbsp;&nbsp;&nbsp;<a href="sdclass.asp?<%=C_ID%>&t=0,0,0">全部分类</a></td>
  </tr>
												<tr>
													<td colspan="4"><%=m_more(0,0,1,0,0,0,1,0,5,11,9,10,8)%></td>
												</tr>
												<tr>
													<td colspan="4"><HR></td>
												</tr>

												<tr>
													<td colspan="4"><%=n_more(0,1,1,0,0,1,1,3,11,9,10,8)%></td>
												</tr>
</table>

	
	<!--开始-->
<div align="center">
<TABLE>
<TR>
<TD width="210">
<TABLE height=335 cellSpacing=0 cellPadding=0 border=0 class="dq1">
        <TBODY>
        <TR>
          <TD align=middle class="dq4" height=30><b>最 新 加 盟</b></TD></TR>
        <TR>
        <TR>
          <TD width=210>
            <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
              <TBODY>
              <TR>
                <TD align=middle height=97 >
<%=dq_more(0,0,0,0,15,3,1,1,13,10,18,0,0,0)%>
                </TD></TR>
</TBODY></TABLE></TD></TR>
</TBODY></TABLE>
      <table width="210" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
      <TABLE cellSpacing=0 cellPadding=0 border=0 height="335" class="dq1">
        <TBODY>
        <TR>
          <TD align=middle class="dq4" width=210 height=30><b>点 击 热 门</b></TD></TR>
        <TR>
        <TR>
          <TD  valign="top">
<%=dq_more(0,0,0,0,15,2,0,2,11,10,18,1,0,0)%></TD></TR>
        
        </TBODY></TABLE></td>
		<td width="1" valign="top"></td>
	<TD width="760" valign="top">
<table cellSpacing="0" cellPadding="0" width="760" border="0"  class="dq1">
			<tr>				
				<td align="middle" width="198"  height="30" class="dq4"><b>店 铺 展 示</b></td>				
			</tr>
			<tr vAlign="top" align="middle">
				<td  colSpan="5"  bgcolor="#EEEEEE">
				<table cellSpacing="0" cellPadding="0" width="100%" border="0">
					<tr>
						<td>
						<table cellSpacing="0" cellPadding="0" width="100%" border="0">
							<tr>
								<td><table width="100%" height="22" border="0" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td height="23" valign="top"> 
                      <table width="99%" height="20" border="0" cellpadding="4" cellspacing="1" bgcolor="#FFFFFF" align="center">
                      <tr onMouseOver="this.style.backgroundColor='#c5edfc';" onMouseOut="this.style.backgroundColor='#E6F7FF'"  class="dq2"> 
                          <td width="30%" height="20" class="F16">&nbsp;<b>单位名称</b></td>
                          <td width="13%" align="center"><b>地区</b></td>
                          <td width="12%" align="center"><b>电话</b></td>
                          <td width="20%" align="center"><b>加入时间</b></td>
                          <td width="6%" align="center"><b>点击</b></td>
                        </tr>
                        <tr>
						  <td width="100%" colspan="5" height="1" background="../images/dotline.gif"></td>
						</tr>
<%
Dim rsqy,sqlqy,qynavigate,TempText
Dim rdsTotalRec 
Dim intPage        '/当前页/
Dim intTotalRec	   '/所有记录/
Dim intInfoCount
Dim p,ij,strpage	       '/定义一些临时变量/
Dim intInfoCot
Dim searchname,theitem,thecity,thepage,currentpage,pagename,totalpage,page,itype,icity,cs,g,b,bb
city_twoid=0
city_threeid=0
type_twoid=0
type_threeid=0
Function   GetCount(S,C)   
          GetCount=0   
          Dim   I   
          Dim   Count  
		  count=0 
          For   I=1   to   Len(S)   
              If   Mid(S,I,1)=C   Then   Count=Count+1   
          Next   
          GetCount=Count   
      End Function   

searchname=HtmlEncode(request("keyword"))
theitem=HtmlEncode(request("item"))
thecity=HtmlEncode(request("icity"))
itype=getcount(theitem,"a")
icity=getcount(thecity,"a")

If searchname<>"" then
  select case itype
  case "0"
  type_oneid=theitem
  case "1"
  i=InStr(theitem,"a")
  type_oneid=left(theitem, i-1)
  ii=len(theitem)
  type_twoid=Right(theitem,ii-i)
  case "2"
  i=InStr(theitem,"a")
  type_oneid=left(theitem, i-1)
  ii=len(theitem)
  theitem=Right(theitem,ii-i)
  i=InStr(theitem,"a")
  type_twoid=left(theitem, i-1)
  ii=len(theitem)
  type_threeid=Right(theitem,ii-i)
  end Select 
End if

 '搜索地区判别,屏蔽的为不定地区,与label.asp搜索条结合======= 
  'select case icity
  'case "0"
  'city_oneid=thecity
  'case "1"
  'i=InStr(thecity,"a")
  'city_oneid=left(thecity, i-1)
  'ii=len(thecity)
  'city_twoid=Right(thecity,ii-i)
  'case "2"
  'i=InStr(thecity,"a")
  'city_oneid=left(thecity, i-1)
  'ii=len(thecity)
  'thecity=Right(thecity,ii-i)
  'i=InStr(thecity,"a")
  'city_twoid=left(thecity, i-1)
  'ii=len(thecity)
  'city_threeid=Right(thecity,ii-i)
  'end Select

If HtmlEncode(request("icity"))<>"总站"  And searchname<>"" Then

cs=split(thecity,"a")
city_oneid=strint(cs(0))
city_twoid=strint(cs(1))
city_threeid=strint(cs(2))
c1=strint(cs(0))
c2=strint(cs(1))
c3=strint(cs(2))

End if
'===================================
    set rsqy = Server.CreateObject("ADODB.Recordset")
    if theitem="全部" and thecity="总站" then'=========1=========
    sqlqy="select * from DNJY_user where sdname Like '%"& searchname &"%' and sdkg=1 order by sdgd "&DN_OrderType&",id "&DN_OrderType&""
    end If'=========1=========
    
	if theitem<>"全部" and thecity="总站" then'=========2=========
	if type_oneid<>0 and type_twoid=0 and type_threeid=0 then 
    sqlqy="select * from DNJY_user where sdname Like '%"& searchname &"%' and type_oneid="&type_oneid&" and sdkg=1 order by sdgd "&DN_OrderType&",id "&DN_OrderType&""
	end if
	if type_oneid<>0 and type_twoid<>0 and type_threeid=0 then 
	sqlqy="select * from DNJY_user where sdname Like '%"& searchname &"%' and type_oneid="&type_oneid&" and type_twoid="&type_twoid&" and sdkg=1 order by sdgd "&DN_OrderType&",id "&DN_OrderType&""
	end if
	if type_oneid<>0 and type_twoid<>0 and type_threeid<>0 then 
	sqlqy="select * from DNJY_user where sdname Like '%"& searchname &"%' and type_oneid="&type_oneid&" and type_twoid="&type_twoid&" and type_threeid="&type_threeid&"and sdkg=1 order by sdgd "&DN_OrderType&",id "&DN_OrderType&""
	end if
	end If'=========2=========
	
	if theitem="全部" and thecity<>"总站" then'=========3=========
	if city_oneid<>0 and city_twoid=0 and city_threeid=0 then 
    sqlqy="select * from DNJY_user where sdname Like '%"& searchname &"%' and city_oneid="&city_oneid&" and sdkg=1 order by sdgd "&DN_OrderType&",id "&DN_OrderType&""
	end if
	if city_oneid<>0 and city_twoid<>0 and city_threeid=0 then 
	sqlqy="select * from DNJY_user where sdname Like '%"& searchname &"%' and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and sdkg=1 order by sdgd "&DN_OrderType&",id "&DN_OrderType&""
	end if
	if city_oneid<>0 and city_twoid<>0 and city_threeid<>0 then 
	sqlqy="select * from DNJY_user where sdname Like '%"& searchname &"%' and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid&" and sdkg=1 order by sdgd "&DN_OrderType&",id "&DN_OrderType&""
	end if
	end If'=========3=========

	if theitem<>"全部" and thecity<>"总站" then'========4=========
	if type_oneid<>0 and type_twoid=0 and type_threeid=0 then
	if city_oneid<>0 and city_twoid=0 and city_threeid=0 then
	sqlqy="select * from DNJY_user where sdname Like '%"& searchname &"%' and type_oneid="&type_oneid&" and city_oneid="&city_oneid&" and sdkg=1 order by sdgd "&DN_OrderType&",id "&DN_OrderType&""
    end if
	if city_oneid<>0 and city_twoid<>0 and city_threeid=0 then
	sqlqy="select * from DNJY_user where sdname Like '%"& searchname &"%' and type_oneid="&type_oneid&" and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and sdkg=1 order by sdgd "&DN_OrderType&",id "&DN_OrderType&""
    end if
	if city_oneid<>0 and city_twoid<>0 and city_threeid<>0 then
	sqlqy="select * from DNJY_user where sdname Like '%"& searchname &"%' and type_oneid="&type_oneid&" and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid&" and sdkg=1 order by sdgd "&DN_OrderType&",id "&DN_OrderType&""
    end if
	end if
	
	if type_oneid<>0 and type_twoid<>0 and type_threeid=0 then
	if city_oneid<>0 and city_twoid=0 and city_threeid=0 then
	sqlqy="select * from DNJY_user where sdname Like '%"& searchname &"%' and type_oneid="&type_oneid&" and type_twoid="&type_twoid&" and city_oneid="&city_oneid&" and sdkg=1 order by sdgd "&DN_OrderType&",id "&DN_OrderType&""
	end if
	if city_oneid<>0 and city_twoid<>0 and city_threeid=0 then
	sqlqy="select * from DNJY_user where sdname Like '%"& searchname &"%' and type_oneid="&type_oneid&" and type_twoid="&type_twoid&" and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and sdkg=1 order by sdgd "&DN_OrderType&",id "&DN_OrderType&""
	end if
	if city_oneid<>0 and city_twoid<>0 and city_threeid<>0 then
	sqlqy="select * from DNJY_user where sdname Like '%"& searchname &"%' and type_oneid="&type_oneid&" and type_twoid="&type_twoid&" and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid&" and sdkg=1 order by sdgd "&DN_OrderType&",id "&DN_OrderType&""
	end if
	end if
	
	if type_oneid<>0 and type_twoid<>0 and type_threeid<>0 then
	if city_oneid<>0 and city_twoid=0 and city_threeid=0 then
	sqlqy="select * from DNJY_user where sdname Like '%"& searchname &"%' and type_oneid="&type_oneid&" and type_twoid="&type_twoid&" and type_threeid="&type_threeid&" and city_oneid="&city_oneid&" and sdkg=1 order by sdgd "&DN_OrderType&",id "&DN_OrderType&""
	end if
	if city_oneid<>0 and city_twoid<>0 and city_threeid=0 then
	sqlqy="select * from DNJY_user where sdname Like '%"& searchname &"%' and type_oneid="&type_oneid&" and type_twoid="&type_twoid&" and type_threeid="&type_threeid&" and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and sdkg=1 order by sdgd "&DN_OrderType&",id "&DN_OrderType&""
	end if
	if city_oneid<>0 and city_twoid<>0 and city_threeid<>0 then
	sqlqy="select * from DNJY_user where sdname Like '%"& searchname &"%' and type_oneid="&type_oneid&" and type_twoid="&type_twoid&" and type_threeid="&type_threeid&" and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid&" and sdkg=1 order by sdgd "&DN_OrderType&",id "&DN_OrderType&""
	end if
	end if
	end if'=======================4============================
 
   	if theitem="" and thecity="" then'=======================5============================
	sqlqy="select * from DNJY_user where sdname<>'' and sdkg=1"&tttt&""&ttccdp&" order by sdgd "&DN_OrderType&",id "&DN_OrderType&""
	end If'=======================5============================

	rsqy.open sqlqy,conn,1,1

thepage=request("pageid")
if isnumeric(thepage)=false then
response.write "<script>alert('参数错误，窗口关闭！');history.back();</script>"
response.end
end if

if rsqy.eof Then
response.write"<tr onMouseOver=""this.style.backgroundColor='#c5edfc';"" onMouseOut=""this.style.backgroundColor='#E6F7FF'""  class=""dq2""><td width=""50%"" height=""20"" class=""F16"" colspan=""5"">"
IF c1>0 Then TempText=conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
response.write "对不起，没有找到符合"&TempText&"条件的网店！"

else

intInfoCot=20'每页条数
intPage = Replace(Trim(Request("Page")),"'","")
intPage = Replace(Trim(Request("Page"))," ","")
if intPage = "" then
	intPage = 1
else
	intPage = Clng(intPage)
end if

intTotalRec =rsqy.recordcount	
if not rsqy.EOF then
	if intTotalRec mod intInfoCot = 0 then
		n = intTotalRec \ intInfoCot
	else
		n = intTotalRec \ intInfoCot + 1
	end if
	rsqy.MoveFirst
	if intPage > n then intPage = n
	if intPage < 1 then intPage = 1
	rsqy.Move (intPage - 1) * intInfoCot
	do while not rsqy.EOF and intInfoCount < Clng(intInfoCot)
	intInfoCount = intInfoCount + 1%>
                        <tr onMouseOver="this.style.backgroundColor='#c5edfc';" onMouseOut="this.style.backgroundColor='#E6F7FF'"  class="dq2"> 
                          <td width="50%" height="20" class="F16"><%if rsqy("vip")=1 then%>&nbsp;<img border="0" src="../images/vip.gif" alt="VIP会员"><%end if%><span style="font-size: 9pt">&nbsp;</span><a href="user/com.asp?com=<%=Cstr(rsqy("username"))%>&<%=T_ID%>&<%=C_ID%>" target="_blank" class="a01"><span style="font-size: 9pt"><% 
response.write Replace (rsqy("sdname"),searchname,"<font color=#FF0000>" & searchname & "</font>") 
%></span></a></td>
                          <td width="13%" align="center"><%qynavigate="<a href=""sdclass.asp?c="&rsqy("city_oneid")&",0,0"">"&rsqy("city_one")&"</a>"
	IF rsqy("city_two")<>"" and not isnull(rsqy("city_two")) Then qynavigate=qynavigate&" / <a href=""sdclass.asp?t="&rsqy("city_oneid")&","&rsqy("city_oneid")&",0"">"&rsqy("city_two")&"</a>"
	IF rsqy("city_three")<>"" and not isnull(rsqy("city_three")) Then qynavigate=qynavigate&" / <a href=""sdclass.asp?t="&rsqy("city_oneid")&","&rsqy("city_twoid")&","&rsqy("city_threeid")&""">"&rsqy("city_three")&"</a>"
	response.write qynavigate%></td>
                         <td width="12%" align="center"><% = rsqy("dianhua")%></td>
                          <td width="11%" align="center"><% =rsqy("dpdata")%></td>
                          <td width="6%" align="center"><% =rsqy("sdhits")%></td>
                        </tr>
                        <tr>
						  <td width="100%" colspan="5" height="1" background="../images/dotline.gif"></td>
						</tr>
<%rsqy.MoveNext
	loop
else%>
						<tr>
						  <td width="100%" colspan="5" height="25">&nbsp;<font color="#FF0000">对不起!没有相匹配的信息</font>...</td>
						</tr>
<%end if

if intPage - 1 mod 10 = 0 then
	p = (intPage - 1) \ 10
else
	p = (intPage - 1) \ 10
end if%>
                      </table>
				    </td>
                  </tr>
                </table>
                <table border="0" cellspacing="0" width="100%" cellpadding="0">
				  <tr>
					<td width="100%" height="30" align="center" valign="top">
	  <table border="0" width="98%" cellspacing="0" cellpadding="0" align="center">
	    <form method="post" action="?<%=T_ID%>&<%=C_ID%>" name="Form2">
	    <input type="hidden" name="KeyWord" value="<% =searchname%>">
	    <input type="hidden" name="item" value="<% =theitem%>">
	    <input type="hidden" name="icity" value="<% =thecity%>">
	    <tr>
		  <td width="38%" height="30" valign="middle">页次：<b><%= intPage %></b>/<b><%=n%></b>页 每页<b><%= intInfoCot %></b>条 共<b><%= intTotalRec %></b>条</td>
		  <td width="62%" height="30" valign="middle"><div align="right">分页：
			<%
			strpage="&KeyWord="&searchname&"&item="&theitem&"&icity="&thecity&"&"&T_ID&"&"&C_ID&""
				if intPage = 1 then
					Response.Write "<font face=webdings>9</font>   "
				else
					Response.Write "<a href='?Page=1"&strpage&"' title=首页><font face=webdings>9</font></a>   "
				end if
				if p * 10 > 0 then Response.Write "<a href='?Page="&Cstr(p*10)&""&strpage&"' title=上十页><font face=webdings>7</font></a>   "
				Response.Write "<b>"
				for ij = p * 10 + 1 to P * 10 + 10
					   if ij = intPage then
				          Response.Write "<font color=""#FF0000"">"+Cstr(ij)+"</font> "
					   else
					      Response.Write "<a href='?Page="&Cstr(ij)&""&strpage&"'>"+Cstr(ij)+"</a> "
					   end if
					if ij = n then exit for
				next
				Response.Write "</b>"
				if ij < n then Response.Write "<a href='?Page="&Cstr(ij)&""&strpage&"' title=下十页><font face=webdings>8</font></a>   "
				if intPage=n then
					Response.Write "<font face=webdings>:</font>   "
				else
					Response.Write "<a href='?Page="&Cstr(n)&""&strpage&"' title=尾页><font face=webdings>:</font></a>   "
				end if
			%>
			转到：<input type="text" name="Page" size="2" maxlength="10" value="<%= intPage %>" class="face"> <input type="submit" value="Go" name="submit" class="button">
			</div></form><%end If%></td>
		</tr>
		
	  </table>
					</td>
				  </tr>
				</table></td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
<table width="760" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
<div class="toleft">
<table  id="Table5" cellSpacing="0" cellPadding="5"  border="0" width="100%"  class="dq1">
						<tr>
							<td>
		<table width="100%" border="0" cellspacing="0" cellpadding="0" >
        <tr> 
          <td class="middle" align="center" valign="top"> 
            <div class="top"><span style="float:left;padding-top:2px;"><%=dq_search(0,1,1)%></div>
        </td>
        </tr>
      </table>

							</td>
						</tr>
					</table>
</div>	
      <table width="760" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
<!---730广告开始----><table width="100%" border="0" cellspacing="0" cellpadding="0" class="dq1">
        <tr> 
          <td class="middle" align="center" valign="top"> 
            <div class="top"> 
<script src="<%=adjs_path("ads/js","syzb730",c1&"_"&c2&"_"&c3)%>"></script>
  </div>
        </td>
        </tr>
</table><!---730结束---->		
		</TD>
</TR>
</TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=980 align=center border=0>
  <TBODY>
  <TR><%conn.execute "UPDATE DNJY_user SET mphits=mphits+1 WHERE username='"&username&"'" '点击计数
   rsqy.close
   set rsqy=nothing%>
    <TD width="769">
    <p align="center"><!--#include file=footer.asp--></TD></TR>                         
  <TR>
<TD width="769"></TD></TR></TBODY></TABLE>
</div>
<!--结束-->

</body>
</html>
