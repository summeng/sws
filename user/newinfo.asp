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
'天天供求信息远程调用代码
'将此调用代码放在任何页面：<script language="javascript" src="网址及路径/user/newinfo.asp?xxsx=1&tt1=0&tt2=0&tt3=0&c1=0&c2=0&c3=0&xxlx=0&xxtj=0&gd=0&tp=0&xxsj=0&jj=0&gdxx=0&ys=0&hits=0&btw=20&s=10&l=1&zt=13&hg=10"></script>

'调用参数说明。依次：xxsx信息属性（1最新、2竞价、3推荐、4热门），tt1一级信息分类（0不分、其它按该类ID号），tt2二级信息分类（0不分、其它按该类ID号），tt3三级信息分类（0不分、其它按该类ID号），c1一级地区分类（0不分、其它按该类ID号），c2二级地区分类（0不分、其它按该类ID号），c3三级地区分类（0不分、其它按该类ID号），xxlx信息类型（0不显1显），xxtj推荐标志（0不显1显），gd固顶标志（0不显1显），tp图片标志（0不显1显），xxsj时间（0不显、1显月日、2显年月日、3显复杂），jj竞价提示（0不显1显），gdxx更多（0不显1显），ys标题颜色（0不显1显），hits点击数（0不显1显），btw标题长度（以字节计），s显示条数，l列数，zt字号，hg行高


'############以下请勿修改！######################
dim TempText,g,iii,i,b,tj,bb,gdxxx,yss,b2,xxsx,tt1,tt2,tt3,c1,c2,c3,xxlx,xxtj,gd,tp,xxsj,jj,gdxx,ys,hits,btw,s,l,zt,hg
xxsx=Trim(Request("xxsx"))
tt1=Trim(Request("tt1"))
tt2=Trim(Request("tt2"))
tt3=Trim(Request("tt3"))
c1=Trim(Request("c1"))
c2=Trim(Request("c2"))
c3=Trim(Request("c3"))
xxlx=Trim(Request("xxlx"))
xxtj=Trim(Request("xxtj"))
gd=Trim(Request("gd"))
tp=Trim(Request("tp"))
xxsj=Trim(Request("xxsj"))
jj=Trim(Request("jj"))
gdxx=Trim(Request("gdxx"))
ys=Trim(Request("ys"))
hits=Trim(Request("hits"))
btw=strint(Trim(Request("btw")))
s=strint(Trim(Request("s")))
l=strint(Trim(Request("l")))
zt=Trim(Request("zt"))
hg=Trim(Request("hg"))

gdxxx=""
If gdxx=1 And xxsx=0 Then gdxxx="&xxsx=0"
If gdxx=1 And xxsx=1 Then gdxxx="&xxsx=1"
If gdxx=1 And xxsx=2 Then gdxxx="&xxsx=2"
If gdxx=1 And xxsx=3 Then gdxxx="&xxsx=3"
If gdxx=1 And xxsx=4 Then gdxxx="&xxsx=4"
If gdxx=1 And xxsx=5 Then gdxxx="&xxsx=5"

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
	IF tt1=0 then
	tttt=""
	elseIF tt3>0 then
	tttt=" and type_oneid="&tt1&" and type_twoid="&tt2&" and type_threeid="&tt3
	elseIF tt2>0 then
	tttt=" and type_oneid="&tt1&" and type_twoid="&tt2
	Else
	tttt=" and type_oneid="&tt1
	End If
Call OpenConn
set rs=server.createobject("ADODB.recordset")
sql = "select * from DNJY_info where yz=1"
if jiaoyi=0 Then sql=sql&" and zt="&jiaoyi
if overdue=0 Then sql=sql&" and dqsj>"&DN_Now&""
sql=sql&""&ttcc
sql=sql&""&tttt
if xxsx=2 Then sql=sql&" and hfxx=1"
if xxsx=3 Then sql=sql&" and tj>0"
if xxsx=4 Then sql=sql&" and llcs>="&hotbz&""

sql=sql&" order by "
if xxsx=2 Then
sql=sql&" hfje/fbts "&DN_OrderType&"" 
elseif xxsx=4 Then
sql=sql&" llcs "&DN_OrderType&""
else
sql=sql&" b "&DN_OrderType&",fbsj "&DN_OrderType&""
end If

Select Case xxsx
	Case 1
	b2="最新"
	Case 2
	b2="竞价"
	Case 3	
	b2="推荐"
	Case 4	
	b2="热门"
	Case Else	
	b2=""
End Select

rs.open sql,conn,1,1%>
TempText=""
<%if rs.eof Then%>
TempText=TempText+"暂无"
<%IF c1>0 Then%>
TempText=TempText+"<%=conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)%>"
<%End if%>
TempText=TempText+"<%=b2%>信息"
<%else
iii=0%>
TempText=TempText+"<table width='100%'  border='0' cellspacing='2' cellpadding='0'>"
<%do while not rs.eof
i=i+1
b=trim(rs("b"))
tj=trim(rs("tj"))
bb=len(b)
IF iii Mod l=0 Then%>
TempText=TempText+"<tr height='<%=hg%>'>"
<%End If%>
TempText=TempText+"<td>"
<%If xxlx=1 then%>
TempText=TempText+"<SPAN style='font-size=<%=zt%>px'>[<%=rs("leixing")%>]</SPAN> "
<%End if%>
<%if rs("tupian")<>"0" And tp=1 Then%>
TempText=TempText+"<img src='http://<%=web%>/images/num/pic.gif' alt='有图片'>" 
<%End if%>
TempText=TempText+"<a target='_blank' href='http://<%=web%>/html/<%=rs("id")%>.html' title='<%=rs("name")%>发布于<%=month(rs("fbsj"))%>月<%=day(rs("fbsj"))%>日'>"
<%if rs("a")<>"0" And ys=1 Then
yss=rs("a")
Else
yss="000000"
End if%>
TempText=TempText+"<SPAN style='color:#<%=yss%>;font-size=<%=zt%>px'><%=CutStr(rs("biaoti"),btw,"...")%></SPAN></a>"
<%if b<>0 And gd=1 then%>
TempText=TempText+"<img src='http://<%=web%>/images/num/jsq.gif' alt='置顶<%=b%>级'>"
<%for g=1 to bb%>
TempText=TempText+"<img src='http://<%=web%>/images/num/<%=Mid(b,g,1)%>.gif' alt='置顶<%=b%>级'>"
<%Next
end If
if tj<>0 And xxtj=1 Then%>
TempText=TempText+"<img src='http://<%=web%>/images/jian.gif' width='15' height='15' alt='本站推荐'>"
<%End If
if rs("hfxx")=1 And rs("dqsj")>now() And jj=1 Then%>
TempText=TempText+"<span title=竞价平均每天<%=FormatNumber(rs("hfje")/rs("fbts"),2,-1)%>元 style='cursor:help'>(<%=FormatNumber(rs("hfje")/rs("fbts"),2,-1)%>)</span>"  
<%End If
IF hits=1 Then%>
TempText=TempText+"<font color='#999999' ><span title=此信息有<%=rs("llcs")%>人关注过 style='cursor:help'>[<%=rs("llcs")%>]</span></font>" 
<%End If
If xxsj>0 Then%>
TempText=TempText+"<td>"
<%End If
If xxsj=1 Then%>
TempText=TempText+"<font color='#999999' ><%=mid(DateValue(rs("fbsj")),6)%></font>" 
<%End If
If xxsj=2 Then%>
TempText=TempText+"<font color='#999999' ><%=mid(DateValue(rs("fbsj")),1)%></font>"
<%End If
If xxsj=3 Then%>
TempText=TempText+"<%=dicksj2(rs("fbsj"))%>"
<%End If
If xxsj>0 Then%>
TempText=TempText+"</td>"
<%End if%>
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
	TempText=TempText+"<tr height='<%=hg%>'><td align='right'><A href='http://<%=web%>/list.asp?c=<%=c1%>,<%=c2%>,<%=c3%><%=gdxxx%>' target=_blank><SPAN style='font-size=<%=zt%>px'>更多>>></SPAN></a></td></tr>"
<%End If%>
TempText=TempText+"</table>"
<%end if
Call CloseRs(rs)
Conn.Close
Set Conn = Nothing%>
document.write (TempText)