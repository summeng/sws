<!--#include file="sdtop.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Dim m%>
<table width="960" align="center" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="190" align="left" valign="top" class="top_m_txt01">
<!--#include file="sdleft.asp"-->
</td>
<td width="5"></td>
<td width="765" valign="top">
<table width="765" border="0" cellpadding="0" cellspacing="0" id="line">
      <tr>
        <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" >
          <tr>		    
            <td><div id="lt2"><span style="margin-left:50px;">公司简介</span></div></td>            
          </tr>
        </table>
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="30%"><div class="aboutpic"><%if rs("logo")<>"" then %> <a target="_blank" title="点击放大&gt;&gt;&gt;" href="../<%=rs("logo")%>"><img border="0" src="../<%=rs("logo")%>" title="点击放大" width="220" height="180"></a><%else%><img src="../images/sdimg/nopic.gif" align="middle" border=0 alt="没有店标"><%end if%></div></td>
              <td width="70%" valign="top"><div style="margin:8px; line-height:23px;;font-size:10pt;"><%=HTMLDecode(CutStr(rs("sdjj"),500,"..."))%> <a href="sdjj.asp?com=<%=com%>">详细>>></a></div></td>
            </tr>
          </table></td>
      </tr>
    </table>
	<table width="100%" border="0" cellpadding="0" cellspacing="0" id="line">
      <tr>
        <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td><div id="lt2"><span style="margin-left:50px;">产品推荐</span></div></td>
          </tr>
        </table>
          <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" style="margin:0px auto">
            <tr>
              <td valign="top"><%=Productstgd(0,0,755,110,18,1)%>
<SCRIPT type=text/javascript>
var speed=30

marquePic2.innerHTML=marquePic1.innerHTML 
function Marquee(){ 
if(demo.scrollLeft>=marquePic1.scrollWidth){ 
demo.scrollLeft=0 
}else{ 
demo.scrollLeft++ 
}} 
var MyMar=setInterval(Marquee,speed) 
demo.onmouseover=function() {clearInterval(MyMar)} 
demo.onmouseout=function() {MyMar=setInterval(Marquee,speed)} 
</SCRIPT>    </td>
            </tr>
          </table></td>
      </tr>
    </table>
	<%If jf<adjf Or adjf=0 then%>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td align="center"><script language=javascript src="<%=adjs_path("ads/js","xxfl1",c1&"_"&c2&"_"&c3)%>"></script></td>
            </tr>
          </table><%End if%>
	<table width="100%" border="0" cellpadding="0" cellspacing="0" id="line">
      <tr>
        <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td><div id="lt2"><span style="margin-left:50px;">最新产品</span></div></td>
          </tr>
        </table>
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" style="margin:5px auto">
            <tr>
              <td><table>
 <%dim Rstj_A,Sqltj_A
 k=0
 i=0
set rstj_a=server.createobject("adodb.recordset")
sqltj_a="select * from DNJY_info where yz=1  and username='"&com&"' order by fbsj "&DN_OrderType&",ID "&DN_OrderType&"" 	      
rstj_a.open sqltj_a,conn,1,1        
Do while Not rstj_a.eof and not rstj_a.bof   
%> 
<td align="center" class=tmpstr valign="top" style="; border:1px solid #CCCCCC; background-color:#F9F9F9" width="25%">
<a title="<%=rstj_a("biaoti")%>" target="_blank" href="../<%=x_path("",rstj_a("id"),"",rstj_a("Readinfo"))%>"><%if rstj_a("tupian")<>"0" and rstj_a("tupian")<>"" then%><IMG src="../UploadFiles/infopic/<%=rstj_a("tupian")%>" width="140" height="110" border=1 style="border: 1px solid #FFFFFF; ; padding-left:2px; padding-right:2px" hspace="0" ><br><u><b><%=CutStr(rstj_a("biaoti"),18,"")%></b></u><font size="1"></font></a><%else%><img src="../images/nophoto2.gif" alt="无图片" width="140" height="110" border="0"><br><u><b><%=CutStr(rstj_a("biaoti"),16,"")%></b></u><font size="1"></font><%end if%></td>
<%     
k=k+1
i=i+1
If i=5 then
response.write "</tr>"
i=0
End If
rstj_a.movenext
if k>=10 then exit do
loop 
rstj_a.close
set rstj_a=nothing    
%> </table></td>
            </tr>
          </table>
		  </td>
      </tr>
    </table>
	<table width="100%" border="0" cellpadding="0" cellspacing="0" id="line">
      <tr>
        <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td><div id="lt2"><span style="margin-left:50px;">新闻中心</span></div></td>
          </tr>
        </table>
          <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" style="margin:10px auto">
            <tr>
              <td valign="top"><%=usernewslist(1,3,15,1,1,28,1,13)%></td>
            </tr>
          </table></td>
      </tr>
    </table>
</tr>
</table>
<!--#include file="sdend.asp"-->
<%function usernewslist(id,lx,ts,qz,tp,btw,dj,rq)

'id类别，lx文章类型（根据ID号定），ts显示条数，qz标题前缀0不显1·2*，tp图片标志，btw标题长度，dj点击数，rq日期1－13种
Dim sqlt,rst,rsnews,ClassName
sqlt="select * from DNJY_userNewsClass where ID="&id&""
set rst=server.createobject("ADODB.Recordset")
rst.open sqlt,conn,1,1
if rst.eof Then
Response.Write "<p><br>暂无分类或分类ID号设置不对！"
rst.close
set rst=nothing
else
ClassName=rst("ClassName")
rst.close
set rst=nothing%>
<div class="shadow2">
 <table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse" class="stock">
  <tr><TD width=300>
<%Dim ae                
ae=0
set rsnews=server.createobject("adodb.recordset")
sql = "select * from DNJY_userNews where username='"&com&"' order by id desc"
rsnews.open sql,conn,3,3
if rsnews.eof Then
Response.Write "<p><br>暂无文章"
else
do while not rsnews.eof%>
     <tr>
       <td valign="bottom"><div class="infoft1"><%if lx=1 then%>[<%=ClassName%>]<%end if%><%If qz=1 then%>·<%elseIf qz=2 then%>*<%End if%><% if rsnews("SavePathFileName")<>"" And tp=1 then response.write "<font color=#0000ff>[图]</font>" end if%>
	<%response.write "<a href=article.asp?id="&Cstr(rsnews("id"))&"&com="&com&" target=_blank title=此条被浏览过"&rsnews("hits")&"次>"
response.write CutStr(trim(rsnews("title")),btw,"...")
response.write "</a>"%><%If dj=1 then%><FONT color=#999999> <%=rsnews("hits")%>|</font><%End if%>
<FONT color=#999999><%=FormatDate(rsnews("addtime"),rq)%></font></td>
     </tr>
     <%
                ae=ae+1
                if ae>=ts then exit do
                rsnews.movenext
                Loop
                End if
                rsnews.close
                set rsnews=Nothing
    End if            
   Response.Write "</TD></tr></table>"
end Function

'*************************************************
'标签名：ProductsGd
'作  用：产品图文(滚动)展示调用（无分页）
'参  数：Num－调用数量0为全部，ls－每行数量0不分行，tw－区域宽度，th－区域高度，btw标题长度,ys标题颜色
'*************************************************
function ProductstGd(Num,ls,tw,th,btw,ys)
Dim Rstj_b,temptext,ProductName
set Rstj_b = server.createobject("adodb.recordset") 
  sql="select * from DNJY_info where yz=1 and tj=1 and username='"&com&"' order by fbsj "&DN_OrderType&",ID "&DN_OrderType&""
  Rstj_b.open sql,conn,1,1
  if Rstj_b.eof Then
  temptext="<td>暂无符合条件产品</td>"  
  else 
 temptext="<DIV id=demo style=""OVERFLOW: hidden; WIDTH:"&tw&"px; align: center"">"
 temptext=temptext& "<TABLE cellSpacing=0 cellPadding=0 align=center border=0>"
 temptext=temptext& "<TBODY><TR><TD id=marquePic1 vAlign=top>"
 temptext=temptext& "<TABLE height="&th&" cellSpacing=0 cellPadding=0 width=660 border=0>"
 temptext=temptext& "<TBODY>"
 temptext=temptext& "<TR><TD align=middle width=103><table border=""0"" cellspacing=""0"" cellpadding=""0""><tr>"
          dim ii
		  i=0
          ii=1
          do while not Rstj_b.eof
ProductName=CutStr(Rstj_b("biaoti"),btw,"...")
If ys=1 Then ProductName="<font color=""#"&Rstj_b("a")&""">"&CutStr(Rstj_b("biaoti"),btw,"...")&"</font>"  
temptext=temptext& "<td><table width=""100%"" border=""0"" cellpadding=""5"" cellspacing=""1"">"
temptext=temptext& "<tr><td width=""104""><a href="""&x_path("",Rstj_b("id"),"",rstj_b("Readinfo"))&""" class=""img"" target=_blank title="""&Rstj_b("biaoti")&""">"
if rstj_b("tupian")<>"0" and rstj_b("tupian")<>"" then
temptext=temptext& "<IMG src=""../UploadFiles/infopic/"&rstj_b("tupian")&""" width=140 height=110 border=1 style=""border: 1px solid #FFFFFF; padding-left:2px; padding-right:2px"" hspace=""0"" >"
Else
temptext=temptext& "<img src=""../images/nophoto2.gif"" alt=""无图片"" width=""140"" height=""110"" border=""0"">"
End if
temptext=temptext& "</a></td></tr>"
temptext=temptext& "<tr><td><div align=""center""><a href="""&x_path("",Rstj_b("id"),"",rstj_b("Readinfo"))&""" target=_blank title="""&Rstj_b("biaoti")&""">"&ProductName&"</a></div></td></tr></table></td>"
 if ii=ls And ls>0 then
 temptext=temptext& "<tr>"
 ii=0
 end If
 i=i+1
 ii=ii+1 
Rstj_b.movenext
if i>=Num And Num>0 then exit do
loop 
temptext=temptext& "</tr></table></TD></TR></TBODY></TABLE></TD><TD id=marquePic2 vAlign=top></TD></TR></TBODY></TABLE></DIV>"
end If
Rstj_b.close
set Rstj_b=nothing
ProductstGd=temptext
end Function%>