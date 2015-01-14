<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
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
<% 
'天天供求信息企业名片远程调用代码
'将此调用代码放在任何页面：<script language="javascript" src="网址及路径/user/newhy.asp?c1=0&c2=0&c3=0&t1=0&t2=0&t3=0&gdxx=0&btw=24&s=10&l=1&zt=13&hg=10"></script>

'调用参数说明。依次：c1一级地区分类（0不分、其它按该类ID号），c2二级地区分类（0不分、其它按该类ID号），c3三级地区分类（0不分、其它按该类ID号）,t1行业一级分类（0不分、其它按该类ID号）,t2行业二级分类（0不分、其它按该类ID号），gdxx更多（0不显1显），btw名称长度（以字节计），s显示条数，l列数，zt字号，hg行高


'############以下请勿修改！######################
dim TempText,iii,i,c1,c2,c3,t1,t2,t3,gdxx,btw,s,l,zt,hg
c1=strint(Trim(Request("c1")))
c2=strint(Trim(Request("c2")))
c3=strint(Trim(Request("c3")))
t1=strint(Trim(Request("t1")))
t2=strint(Trim(Request("t2")))
t3=strint(Trim(Request("t3")))
gdxx=Trim(Request("gdxx"))
btw=strint(Trim(Request("btw")))
s=strint(Trim(Request("s")))
l=strint(Trim(Request("l")))
zt=Trim(Request("zt"))
hg=Trim(Request("hg"))

    Dim ttcc
	ttcc=""
    IF c3>0 Then
	ttcc=" and city_oneid="&c1&" and city_twoid="&c2&" and city_threeid="&c3
	ElseIF c2>0 Then
	ttcc=" and city_oneid="&c1&" and city_twoid="&c2
	ElseIF c1>0 Then
	ttcc=" and city_oneid="&c1
	End If

    Dim tttt
	IF t1=0 then
	tttt=""
	elseIF t3>0 then
	tttt=" and type_oneid="&t1&" and type_twoid="&t2&" and type_threeid="&t3
	elseIF t2>0 then
	tttt=" and type_oneid="&t1&" and type_twoid="&t2
	Else
	tttt=" and type_oneid="&t1
	End If
Call OpenConn
set rs=server.createobject("ADODB.recordset")
sql="select * from [DNJY_user] where mpname<>'' and mpkg=1"
sql=sql&""&ttcc
sql=sql&""&tttt
sql=sql&" order by id "&DN_OrderType&""
rs.open sql,conn,1,1%>
TempText=""
<%if rs.eof Then%>
TempText=TempText+"暂无"
<%IF c1>0 Then%>
TempText=TempText+" <%=conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)%>"
<%End if%>
<%IF t1>0 Then%>
TempText=TempText+" <%=conn.Execute("Select name From DNJY_hy_type Where id="&t1&" and twoid="&t2&" and threeid="&t3&"")(0)%>类 "
<%End if%>
TempText=TempText+"名片"
<%else
iii=0%>
TempText=TempText+"<table width='100%'  border='0' cellspacing='2' cellpadding='0'>"
<%do while not rs.eof
i=i+1
IF iii Mod l=0 Then%>
TempText=TempText+"<tr height='<%=hg%>'>"
<%End If%>
TempText=TempText+"<td>"
TempText=TempText+"<a target='_blank' href='http://<%=web%>/company.asp?username=<%=rs("username")%>'>"
TempText=TempText+"<SPAN style='font-size=<%=zt%>px'><%=CutStr(rs("mpname"),btw,"...")%></SPAN></a>"
TempText=TempText+"</TD>"
<%iii=iii+1
IF iii Mod l=0 then%>
TempText=TempText+"</tr>"
<%End If
if i>=s then exit Do
	Rs.movenext
	Loop	
	IF iii Mod l<>0 then
	Do While not iii Mod l=0%>
	TempText=TempText+"<td></td>"
	<%iii=iii+1
	Loop%>
	TempText=TempText+"</tr>"
	<%End IF%>	
	<%If gdxx=1 Then%>
	TempText=TempText+"<tr height='<%=hg%>'><td align='right'><A href='http://<%=web%>/qyindex.asp?c=<%=c1%>,<%=c2%>,<%=c3%>&t=<%=t1%>,<%=t2%>,<%=t3%>' target=_blank><SPAN style='font-size=<%=zt%>px'>更多>>></SPAN></a></td></tr>"
<%End If%>
TempText=TempText+"</table>"
<%end if
Call CloseRs(rs)
Call CloseDB(conn)%>
document.write (TempText)