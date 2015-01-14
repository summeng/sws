<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/Function.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<%
Call OpenConn
Dim action,selectkey,issue,editionname,Releasetime,Releasetimea,issuea,thumbnail,largerpic,Turnpage,rs_a,sql_a,rs_b,sql_b,rs_c,issueid
'定义地区分类有关开始=============
Dim city_oneid,city_twoid,city_threeid,city_one,city_two,city_three,c,c1,c2,c3,c4,WebSetting,cityweb,i,titleinfo,description,keywords,Cellular_phone,tel,fax,mymail,serve,yzbm,qq,qq_name,cityCompany
c1=0
c2=0
c3=0
c1=city_oneid
c2=city_twoid
c3=city_threeid
c4=""&c1&","&c2&","&c3&""

c=trim(request("c"))
If trim(request("c"))="" Or trim(request("c"))="0" Then c=c4
c=split(c,",")
If c(0)="" Then c(0)=0
c1=strint(c(0))
If ubound(c)<1 Then
c2=0
else
c2=strint(c(1))
End if
If ubound(c)<2 Then
c3=0
else
c3=strint(c(2))
End If
city_oneid=strint(c1)
city_twoid=strint(c2)
city_threeid=strint(c3)

set rs=server.createobject("adodb.recordset")
if city_oneid>0 then
rs.open "select city from DNJY_city where twoid=0 and threeid=0 and id="&city_oneid,conn,1,1
city_one=rs("city")
rs.close
end if
if city_twoid>0 then
rs.open "select city from DNJY_city where id="&city_oneid&" and threeid=0 and twoid="&city_twoid,conn,1,1
city_two=rs("city")
rs.close
end if
if city_twoid>0 and city_threeid>0 then
rs.open "select city from DNJY_city where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid,conn,1,1
city_three=rs("city")
rs.close
end If


If city_oneid>0 Then'分站访问时使用分站
	Sql="Select city,WebSetting,dname From [DNJY_city] Where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid
	Set Rs=Server.CreateObject("ADODB.recordset")
	Rs.open Sql,Conn,1,1
	If Rs.eof Then
	    Call CloseRs(rs)		
		Response.Write("没有地区分类或参数错误，无法正常显示！")
		Response.End
	Else	
		WebSetting = Rs(1)
		cityweb=Rs(2)
	End If
	Call CloseRs(rs)
If Len(WebSetting)<37 or Isnull(WebSetting) Then'当字段为空或虽有分隔符但未设置数据时（指分站）
WebSetting=""'当分站未设置任何数据时清空分站数据准备取总站数据
	For I=0 To 12'到最后需要的项目为止
		'If i<>4  And i<>6 Then WebSetting = WebSetting&Application(CacheName&"_WebSetting")(i) &"∮∮∮" '过滤不需要的
		WebSetting = WebSetting&Application(CacheName&"_WebSetting")(i) &"∮∮∮" '当分站数据为空时使用总站的，并在总站各段数据后加上∮∮∮（除了与分站不对应的项目）
	Next
	WebSetting =Left(WebSetting,Len(WebSetting)-3)'除掉最后三个∮∮∮
End If
WebSetting=Split(WebSetting,"∮∮∮")
title=WebSetting(0)
titleinfo=WebSetting(1)
description=WebSetting(2)
mymail=WebSetting(4)
Tel=WebSetting(5)
Cellular_phone=WebSetting(6)
fax=WebSetting(7)
serve=WebSetting(8)
yzbm=WebSetting(9)
qq=WebSetting(10)
qq_name=WebSetting(11)
cityCompany=WebSetting(12)
Call CloseRs(rs)
Else
Set Rs=Server.CreateObject("ADODB.recordset")
Sql="Select WebSetting From DNJY_config"
rs.open sql,conn,1,1
WebSetting = Rs("WebSetting")
WebSetting=Split(WebSetting,"|")
title=WebSetting(0)
mymail=WebSetting(4)
Tel=WebSetting(5)
Cellular_phone=WebSetting(6)
fax=WebSetting(7)
serve=WebSetting(8)
yzbm=WebSetting(9)
qq=WebSetting(10)
qq_name=WebSetting(11)
Call CloseRs(rs)
End If

 '判断是否有分站LOGO文件存在，否则总站LOGO＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
Function l_path(p,logo)
	Dim TempText,Fso
	TempText=Server.MapPath(p)
	Set Fso = CreateObject("Scripting.FileSystemObject")
	IF Fso.FileExists(TempText&"/"&logo&".jpg") and shows10=1 Then	
	l_path=p&"/"&logo&".jpg"	
	Else
	l_path=p&"/0_0_0.jpg"
	End IF
	Set Fso =Nothing
End Function
 '判断是否分站显示广告和是否有广告JS文件存在，否则显示总站广告＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
Function adjs_path(p,adid,jsid)
	Dim TempText,Fso
	TempText=Server.MapPath(p)
	Set Fso = CreateObject("Scripting.FileSystemObject")
	IF Fso.FileExists(TempText&"/"&adid&"_"&jsid&".js") and shows8=1 Then	
	adjs_path=p&"/"&adid&"_"&jsid&".js"		
	Else
	adjs_path=p&"/"&adid&"_0_0_0.js"	
	End IF
	Set Fso =Nothing
End Function
'定义地区分类有关END===============================

set rs_a=server.CreateObject("ADODB.Recordset")
if request("issue")<>"" And city_oneid>0 then
sql_a="select top 1 * from DNJY_zz_issue where issue="&request("issue")&" and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid&" order by Releasetime"&DN_OrderType&""
ElseIf city_oneid>0 then
sql_a="select * from DNJY_zz_issue where city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid&" order by Releasetime"&DN_OrderType&""
Else
sql_a="select top 1 * from DNJY_zz_issue order by Releasetime"&DN_OrderType&""
end if
rs_a.open sql_a,conn,1,1
if not rs_a.eof or not rs_a.bof then
Releasetimea=rs_a("Releasetime")
issuea=rs_a("issue")
end if
rs_a.close  
set rs_a=nothing 
%>
<HTML><HEAD><TITLE><%=title%>－电子杂志 <%=city_one&city_two&city_three%>版</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<META http-equiv="Content-Language" content="zh-cn">
<META content="<%=title%>,电子杂志,网上杂志,网上DM广告,电子广告">
<META content="<%=title%>,电子杂志,网上杂志,网上DM广告,电子广告,网上看报.网站报刊,电子报刊">
<SCRIPT language=javascript>
function xingqi(rqsj){
<!--
var today = new Date(Date.parse(rqsj.replace(/-/g, '/')));
var day;
if(today.getDay()==0) day = "星期日";
if(today.getDay()==1) day = "星期一";
if(today.getDay()==2) day = "星期二";
if(today.getDay()==3) day = "星期三";
if(today.getDay()==4) day = "星期四";
if(today.getDay()==5) day = "星期五";
if(today.getDay()==6) day = "星期六";
return day;
-->
}
function xiangxia(){
	window.gdt.scrollTop=window.gdt.scrollTop+6
	//window.aac.value=window.gdt.scrollTop+"＝"+window.gdt.scrollHeight
	if (window.gdt.scrollTop<window.gdt.scrollHeight-330&&window.gdt.scrollTop<sp) xiangxia();
	}
function xiangxia1(){
	window.gdt.scrollTop=window.gdt.scrollTop+24
	//window.aac.value=window.gdt.scrollTop+"＝"+window.gdt.scrollHeight
	if (window.gdt.scrollTop<window.gdt.scrollHeight-330&&window.gdt.scrollTop<sp) xiangxia1();
	}
function hg_a(){
		sp=window.gdt.scrollTop+200
		if (sp>window.gdt.scrollHeight-415) sp=window.gdt.scrollHeight-415
		xiangxia();
	}
function xiangshang(){
	window.gdt.scrollTop=window.gdt.scrollTop-6
	//window.aac.value=window.gdt.scrollTop+"＝"+window.gdt.scrollHeight
	if (window.gdt.scrollTop>0&&window.gdt.scrollTop>ps) xiangshang();
	}
function hg_b(){
		ps=window.gdt.scrollTop-200
		if (ps<0) ps=0
		xiangshang();
	}
</SCRIPT>
<script>
function loadThreadFollow(t_id,b_id){
	var targetImg =eval("document.all.followImg" + t_id);
	var targetDiv =eval("document.all.follow" + t_id);
	
	if ("object"==typeof(targetImg)){
		if (targetDiv.style.display!='block'){
			targetDiv.style.display="block";
			
		}else{
			targetDiv.style.display="none";
		}
	}
}
</script>
<STYLE>
A:active {
	COLOR: #222222; TEXT-DECORATION: underline
}
A:link {
	COLOR: #222222; TEXT-DECORATION: none
}
A:visited {
	COLOR: #222222; TEXT-DECORATION: none
}
A:hover {
	COLOR: #222222; TEXT-DECORATION: underline
}
.STYLE2 {font-size: 12px}
</STYLE>

<META content="MSHTML 6.00.2900.3020" name=GENERATOR></HEAD>
<BODY onselectstart="return false;" oncontextmenu="return false" onselectstart="return false"  bottomMargin=0 bgColor=#E3EBF9 leftMargin=0 
topMargin=0 rightMargin=0>
<DIV align=center>
<TABLE height=582 cellSpacing=0 cellPadding=0 width=1000 border=0>
<TBODY>
<TR>
<TD width=25 rowSpan=4></TD>
<TD width=385 colSpan=2 height=20></TD>
<TD width=590 rowSpan=4>
<TABLE height="100%" cellSpacing=0 cellPadding=0 width="100%" border=0>
<TBODY>
<TR>
<TD vAlign=bottom width=36 height=71><IMG src="img/t01.jpg" border=0></TD>
<TD vAlign=bottom width=191 height=71><A href="http://<%=web%>"><IMG alt=<%=title%>－电子杂志 src="<%=l_path("../UploadFiles/logo",c1&"_"&c2&"_"&c3)%>" alt="<%=title%>" border=0></A></TD>
<TD vAlign=bottom width=10 height=71>　</TD>
<TD vAlign=bottom width=350 height=80> <script src="<%=adjs_path("../ads/js","magazine",c1&"_"&c2&"_"&c3)%>"></script>
</TD>
<TD vAlign=bottom width=25 height=71>　</TD></TR>
<TR>
<TD align=middle colSpan=5 height=23>
<TABLE style="FONT-SIZE: 10pt; COLOR: #b08c48" height="100%" cellSpacing=0 cellPadding=0 width="100%" border=0>
<TBODY>
<TR>
<TD width=36 background=img/menubg.jpg></TD>
<TD background=img/menubg.jpg><B>&nbsp; <%=city_one&city_two&city_three%>版 第<%=issuea%>期</B></TD>
<TD align=right width=220 background=img/menubg.jpg><%=formatdatetime(Releasetimea,1)%><SCRIPT>document.write (xingqi("<%=Releasetimea%>"))</SCRIPT>发布</TD>
<TD width=25></TD></TR></TBODY></TABLE></TD></TR>
<TR>
<TD align=middle colSpan=5>
<TABLE height="100%" cellSpacing=0 cellPadding=0 width="100%" border=0>
<TBODY>
<TR>
<TD width=36 rowSpan=2></TD>
<TD height=7></TD>
<TD width=25 rowSpan=2></TD></TR>
<TR>
<TD>
<TABLE height="100%" cellSpacing=0 cellPadding=0 width="100%" border=0>
<TBODY>
<TR>
<TD vAlign=top width=150>
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
<TBODY>
<TR>
<TD style="FONT-SIZE: 14px; COLOR: #ffffff" align=middle bgColor=#799AE1 height=24>·版面导航·</TD></TR>
<TR>
<TD height=436>
<TABLE height="100%" cellSpacing=0 cellPadding=0 width="100%" border=0>
<TBODY>
<TR>
<TD height=9><IMG onmousedown=hg_b(); style="CURSOR: hand" height=9 src="img/s.jpg" width=150 border=0></TD></TR>
<TR>
<TD vAlign=top bgColor=#ffffff style="font-size: 10pt">
<DIV id=gdt style="OVERFLOW-X: hidden; OVERFLOW: hidden; WIDTH: 150px; HEIGHT: 415px">
<DIV class="bmml style01" style="WIDTH: 185px; HEIGHT: 415px">
<TABLE style="FONT-SIZE: 10pt" cellSpacing=0 cellPadding=0 width=150>
<TBODY>
<TR>
<TD height=2></TD></TR>                                
<%
issue=trim(request("issue"))
set rs=server.createobject("adodb.recordset")
if issue<>"" and city_oneid<>"" then
sql= "select * from DNJY_zz_edition where state="&DN_True&" and issue="&issue&" and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid&"  order by id asc,Releasetime"&DN_OrderType&""
ElseIf city_oneid>0 then
sql= "select * from DNJY_zz_edition where state="&DN_True&" and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid&" and issue=(select max(issue) from DNJY_zz_edition)  order by id asc,Releasetime"&DN_OrderType&""
Else
sql= "select top 1 * from DNJY_zz_edition where state="&DN_True&" order by id asc,Releasetime"&DN_OrderType&""
end if
rs.open sql,conn,1,1
if rs.bof or rs.eof then
response.write"暂无版面"
else
if request("Turnpage")="" then
Turnpage=rs("id")
issue=rs("issue")
else
Turnpage=int(trim(request("Turnpage")))
Response.Cookies("issueid")=Turnpage
issue=int(trim(request("issue")))
end if
do while not rs.eof
%>
<TR><TD align=left height=24 <%if rs("id")=Turnpage then
if Turnpage<>"" then
issueid=Turnpage
Response.Cookies("issueid")=Turnpage
else
issueid=rs("id")
Response.Cookies("issueid")=issueid
end if
%>bgColor=#B7CDFC<%end if%>>▲ <A href="?Turnpage=<%=rs("id")%>&issue=<%=rs("issue")%>&c=<%=rs("city_oneid")%>,<%=rs("city_twoid")%>,<%=rs("city_threeid")%>"><%=rs("editionname")%></A></TD></TR>
<%
rs.movenext
loop
end if
Call CloseRs(rs)
%>
<TR>
<TD height=2></TD></TR></TBODY></TABLE></DIV></DIV></TD></TR>
<TR>
<TD height=9><IMG onmousedown=hg_a(); style="CURSOR: hand" height=9 src="img/x.jpg" width=150 border=0></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD>
<TD width=24></TD>
<TD vAlign=top width=94>
<TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
<TBODY>
<TR>
<TD style="FONT-SIZE: 14px; COLOR: #ffffff" align=middle bgColor=#799AE1 height=24>·近期杂志·</TD></TR>
<TR>
<TD vAlign=top bgColor=#ffffff height=436 style="font-size: 10pt" align="center">
<TABLE style="FONT-SIZE: 10pt" cellSpacing=0cellPadding=0 width=94>
<TBODY>
<%
set rs_b=server.createobject("adodb.recordset")
sql_b= "select top 16 * from DNJY_zz_issue where state="&DN_True&" and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid&" order by id"&DN_OrderType&",Releasetime"&DN_OrderType&""
rs_b.open sql_b,conn,1,1
if rs_b.bof or rs_b.eof then
response.write"暂无记录！"
else
do while not rs_b.eof
%>
<TR><TD align=middle <%if rs_b("issue")=issue then%>bgColor=#B7CDFC<%end if%> height=24><A href="?issue=<%=rs_b("issue")%>&c=<%=rs_b("city_oneid")%>,<%=rs_b("city_twoid")%>,<%=rs_b("city_threeid")%>">第<%=rs_b("issue")%>期</A></TD></TR>
<%
rs_b.movenext
loop
end if
rs_b.close
set rs_b=nothing
%>
</TBODY></TABLE>
</TD></TR></TBODY></TABLE></TD>
<TD width=24></TD>
<TD vAlign=top>
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
<TBODY>
<TR>
<TD style="FONT-SIZE: 14px; COLOR: #ffffff" align=middle bgColor=#799AE1 height=24>·各地杂志·</TD>
</TR>
<TR>
<TD  vAlign=top align=middle background=img/t04.jpg height=200>
<TABLE width=250><TBODY>
<TR><TD align=middle height=25 >
<%set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_city] where magazine="&DN_True&" order by id"&DN_OrderType&"" 
rs.open sql,conn,1,1
if rs.eof and rs.bof Then
response.write "没有开办杂志的城市"
Else
%>
<center>
<table width="98%" style="border-collapse: collapse;FONT-SIZE: 13px;" cellpadding="0" border="0" >
<%Dim rs_dq,k,j,page,pages
k=0
j=0
rs.pagesize=25
page=clng(request("page"))
if page=0 then page=1 
pages=rs.pagecount
if page > pages then page=pages
rs.absolutepage=page  

	do while not rs.eof And j<rs.pagesize
    set rs_dq=conn.execute("select count(id) from DNJY_zz_issue where city_oneid="&rs("id")&" and city_twoid="&rs("twoid")&" and city_threeid="&rs("threeid")&"")
	%> <tr>
		<td ><p align="left"><span id="followImg<%=k%>" style="CURSOR: hand" onclick="loadThreadFollow(<%=k%>,5)"><IMG src="img/ja.gif"  border=0><%=conn.Execute("Select city From DNJY_city Where id="&rs("id")&" And twoid=0 And threeid=0")(0)%> <%If rs("twoid")>0 Then response.write conn.Execute("Select city From DNJY_city Where id="&rs("id")&" And twoid="&rs("twoid")&" And threeid=0")(0)%> <%If rs("threeid")>0 Then response.write conn.Execute("Select city From DNJY_city Where id="&rs("id")&" And twoid="&rs("twoid")&" And threeid="&rs("threeid")&"")(0)%></span></td>
		</tr><tr style="display:none" id="follow<%=k%>">
			<td bgcolor="#ffffff"><div ><center>
				<table cellpadding=0 width="100%" border="0" style="border-collapse: collapse" bordercolor="#75BB2C" >
				<%i=1
               set rs_b = Server.CreateObject("ADODB.RecordSet")
               sql_b="select * from [DNJY_zz_issue] where city_oneid="&rs("id")&" and city_twoid="&rs("twoid")&" and city_threeid="&rs("threeid")&" order by issue"&DN_OrderType&",id"&DN_OrderType&""
               rs_b.open sql_b,conn,1,1
               if rs_b.eof then
               response.write "<tr><td style=""border-collapse: collapse;FONT-SIZE: 12px;""><font color=""#ff0000""><p align=""center"">此地区没有发行期刊,请在后台添加！</font></td></tr>"
               else
               %> <tr><%do while not rs_b.eof%>
			   <td height=20 style="border-collapse: collapse;FONT-SIZE: 12px;" ><p>&nbsp;&nbsp;&nbsp;<%If rs_b("state")=true then%><A href="magazine_show.asp?id=<%=rs_b("id")%>&issue=<%=rs_b("issue")%>&c=<%=rs_b("city_oneid")%>,<%=rs_b("city_twoid")%>,<%=rs_b("city_threeid")%>"><%End if%><font color="#666666">第<%=rs_b("issue")%>期</font></a><%If rs_b("state")=false then%>[暂停]<%End if%></td></tr>
			   <%
               i=i+1
			   rs_b.movenext
			   loop
               end if
               rs_b.close
               set rs_b=nothing
               %> </table></center></div></td></tr><%
             k=k+1
			 j=j+1
			 rs.movenext
			 loop
			 %> </table>
			 
	<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" style="border-collapse: collapse;FONT-SIZE: 13px;" >
  <tr bgcolor="#ffffff">   
      <td height="30" align="center"> 
    <%if page<2 then      
    response.write "首页 上一页&nbsp;"
  else
    response.write "<a href=?page=1&c="&city_oneid&","&city_twoid&","&city_threeid&">首页</a>&nbsp;"
    response.write "<a href=?page=" & page-1 & "&c="&city_oneid&","&city_twoid&","&city_threeid&">上一页</a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "下一页 尾页"
  else
    response.write "<a href=?page=" & (page+1) & "&c="&city_oneid&","&city_twoid&","&city_threeid&">"
    response.write "下一页</a> <a href=?page="&rs.pagecount&"&c="&city_oneid&","&city_twoid&","&city_threeid&">尾页</a>"
  end if
   response.write "<br>页次：<strong><font color=red>"&page&"</font>/"&rs.pagecount&"</strong>页 "
    response.write "&nbsp;共<b><font color='#ff0000'>"&rs.recordcount&"</font></b>条记录 <b>"    
%>
      </td>
  </tr>
</table>		 
<%End If%>
</div></TD>
</TR>

</TBODY></TABLE></TD></TR>
</TBODY></TABLE>
</TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR>
<TR>
<TD align=middle width=377 bgColor=#ffffff height=531>
<TABLE cellSpacing=0 cellPadding=0 width=360 border=0>
<TBODY>
<TR>
<TD class=style01 align=middle><A title=查看大图 href="show_largerpic.asp?issue=<%=issue%>&id=<%=issueid%>&city_one=<%=city_one%>&city_two=<%=city_two%>&city_three=<%=city_three%>" target=_blank><IMG src="show_thumbnail.asp?issueid=<%=issueid%>" width=345 border=0 galleryimg="no"  width=370  height=520></A></TD></TR></TBODY></TABLE></TD>
<TD vAlign=top width=8 rowSpan=2><IMG src="img/t02.jpg" border=0></TD></TR>
<TR>
<TD width=377 height=8><IMG src="img/t03.jpg" border=0></TD></TR>
<TR>
<TD width=385 colSpan=2 height=24>
<TABLE style="FONT-SIZE: 13px; COLOR: #222222" cellSpacing=0 cellPadding=0 width="100%" border=0>
<TBODY>
<TR>
<TD align=middle width=100></TD>
<TD align=middle width=218>
<%If issueid<>"" and issue<>"" Then
sql="select top 1 * from DNJY_zz_edition where state="&DN_True&" And id>"&issueid&" and issue="&issue&" and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid&" order by id asc" 
set rs_c=conn.execute(sql) 
if rs_c.eof then 
response.write "<IMG src=img/l.gif  border=0> 已到封面" 
else 
response.write "<a href=?Turnpage="&rs_c("id")&"&issue="&issue&"&c="&city_oneid&","&city_twoid&","&city_threeid&"><IMG src=img/l.gif  border=0> 上一版</a>" 
end if%>
&nbsp;&nbsp;
<%sql="select top 1 * from DNJY_zz_edition where state="&DN_True&" and id<"&issueid&" and issue="&issue&" and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid&" order by id"&DN_OrderType&"" 
set rs_c=conn.execute(sql) 
if rs_c.eof then 
response.write "已到封底 <IMG src=img/r.gif  border=0>" 
else 
response.write "<a href=?Turnpage="&rs_c("id")&"&issue="&issue&"&c="&city_oneid&","&city_twoid&","&city_threeid&">下一版 <IMG src=img/r.gif  border=0></a>"
End if
Else
response.write "暂无版面"
end if%>
<TD align=middle width=100>点看大图</TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></DIV>
<DIV align=center>
<TABLE style="FONT-SIZE: 13px; COLOR: #222222" cellSpacing=0 cellPadding=0  width=1000 border=0>
<TBODY>
<TR>
<TD width=25></TD>
<TD bgColor=#B7CDFC height=3></TD>
<TD width=25></TD></TR>

<TR><TD width=25></TD>
<TD align=middle height=30>Copyright &copy; 2012-<%=CStr(Year(Now()))%> <A style="COLOR: #013162; TEXT-DECORATION: none" href="http://<%=web%>" target=_blank><%=title%></A> All Rights Reserved</TD>
<TD width=25> </TD></TR>
<TR><TD width=25></TD>
<TD align=middle height=30>	<%if Cellular_phone<>"" then%>手机:<%=Cellular_phone%><%end if%> &nbsp; <%if tel<>"" then%>电话:<%=tel%><%end if%><%if fax<>"" then%>    传真:<%=fax%><%end if%><%if mymail<>"" then%>    
<%dim ttemail,ttmail,fqq,fqq_name,n
fqq=qq
fqq_name=qq_name
	ttemail=Replace(""&mymail&"","@","&")
ttmail=Replace(""&mymail&"","@","<IMG SRC=../images/@.gif BORDER=0 width=5 height=8>")%> 客服信箱：<a href="javascript://" onClick="sendmail('<%=ttemail%>?subject=我的建设和意见');return false"><%=ttmail%></a><%end if%>    
<%'客服QQ此文件不需要作任何修改
if fqq<>"" Then 
fqq=replace(fqq,"，",",")
if isnull(fqq_name) or fqq_name="" then fqq_name=fqq
fqq_name=replace(fqq_name,"，",",")
fQQ_NAME=split(fqq_name,",")
fQQ=split(fqq,",")
%> 服务QQ：<%for N=0 to UBound(fQQ)%>
<a class='c' target=blank href='tencent://message/?uin=<%=fQQ(n)%>&Site=<%=title%>&Menu=yes'><%=fQQ(n)%></a>
</script >
<%next%><%end if%>     <%if serve<>"" then%><br>办公地址：<%=serve%><%end if%>  <%if yzbm<>"" then%> 邮政编码：<%=yzbm%><%end if%> <a href="http://www.miibeian.gov.cn" target=_blank><%=icp%></a> </TD>
<TD width=25> </TD></TR>
</TBODY></TABLE></DIV>
</BODY></HTML>
