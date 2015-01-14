<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../Include/Function.asp"-->

<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("12")
dim action,id,dbmname,city_oneid,city_twoid,city_threeid,link,i,paixu,dd
id=strint(request.QueryString("id"))
action=request.querystring("action")
dbmname=trim(request("bmname"))
city_oneid=strint(request.QueryString("city_oneid"))
city_twoid=strint(request.QueryString("city_twoid"))
city_threeid=strint(request.QueryString("city_threeid"))
link=HtmlEncode(request("link"))
paixu=trim(request("paixu"))
If paixu="" Then paixu=0
Call OpenConn
select case action
case "add" 
set Rs=server.CreateObject("adodb.recordset")
Rs.open "select bmname,link,city_oneid,city_twoid,city_threeid,paixu from DNJY_bmcs",conn,1,3
Rs.AddNew
Rs(0)=dbmname
Rs(1)=link
Rs(2)=city_oneid
Rs(3)=city_twoid
Rs(4)=city_threeid
Rs(5)=paixu
Rs.Update
Rs.Close:set Rs=nothing

case "edit"
set Rs=server.CreateObject("adodb.recordset")
Rs.open "select * from DNJY_bmcs where id="&id,conn,1,3
Rs("bmname")=dbmname
Rs("link")=Link
rs("paixu")=trim(request("paixu"))
Rs.Update
Call CloseRs(rs)

case "del"
conn.execute ("delete from DNJY_bmcs where id="&id)
end select
%>
<HTML>
<HEAD>
<TITLE>便民服务管理</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK rel="stylesheet" href="../css/style.css" type="text/css">
</HEAD>
<BODY>
<TABLE width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
  <TR> 
    <TD height="20" align="center" bgcolor="#799AE1"><FONT color="#FFFFFF" style="font-size:14px">&nbsp;
    </FONT><FONT color="#FFFFFF" style="font-size:14px">便 民 服 务 管 理</FONT></TD>
  </TR>
  <TR bgcolor="#FFFFFF">
    <TD height="30" align="center"><TABLE width="100%"  border="0" cellspacing="0" cellpadding="0">
      <TR>
<%set Rs = server.createobject("adodb.recordset")                                            
sql = "select id,twoid,threeid,city from DNJY_city order by id,twoid,threeid"                                            
Rs.open sql,conn,1,1
if rs.eof or rs.bof then
response.write "没有地区分类，无法添加便民服务！请先在[网站系统管理－地区分类]中设置好分类！"
Call CloseRs(rs)
response.End
End if	%>
        <TD width="20%" height="30" bgcolor="#FFFFFF">&nbsp;请选择添加便民服务所属城市：</TD>
        <TD width="30%" bgcolor="#FFFFFF">
          <FONT color="#FFFFFF" style="font-size:14px">

          <SELECT size="1" name="item" onChange="location=this.options[this.selectedIndex].value">
            <OPTION value="?" <%IF city_oneid=0 Then%>selected<%End IF%>>总站</OPTION>
            
            <%do while not Rs.eof
			 if Rs("id")>0 and Rs("twoid")=0 and Rs("threeid")=0 then%>
            <OPTION style="background-color:#F8F4F5 ;color: #FF0000" value="?city_oneid=<%=Rs(0)%>" <%if Rs(0)=city_oneid and city_twoid=0 then%>selected<%end if%>><%=Rs(3)%></OPTION>
            <%elseif Rs(0)>0 and Rs(1)>0 and Rs(2)=0 then%>
            <OPTION style="background-color:#F8F4F5 ;color: #0000FF" value="?city_oneid=<%=Rs(0)%>&city_twoid=<%=Rs(1)%>" <%if Rs(0)=city_oneid and Rs(1)=city_twoid and 0=city_threeid then%>selected<%end if%>>├<%=Rs(3)%></OPTION>
            <%elseif Rs(0)>0 and Rs(1)>0 and Rs(2)>0 then%>
            <OPTION value="?city_oneid=<%=Rs(0)%>&city_twoid=<%=Rs(1)%>&city_threeid=<%=Rs(2)%>" <%if Rs(0)=city_oneid and Rs(1)=city_twoid and Rs(2)=city_threeid then%>selected<%end if%>>├├<%=Rs(3)%></OPTION>
            <%end if
		Rs.movenext                                      
		  Loop

	  
		  Call CloseRs(rs)%>
          </SELECT>
          </FONT>        </TD><TD width="50%">在需要调用的地方调用&lt;iframe marginwidth="0" marginheight="0" src="bmcslist.asp" frameborder="0" width="210" height="200" scrolling="No"  align="center"></iframe>即可。</td>
      </TR>
    </TABLE></TD>
  </TR>
  <TR> 
    <TD bgcolor="#FFFFFF"> <BR>
	<TABLE width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
        <TR align="center" bgcolor="#FFFFFF" height="20"> 
          <TD width="30">编号</TD>
          <TD>网站名称</TD>
          <TD>网址链接</TD>
		  <td>排序<br>(数值大为先)</td>
		  <td>隶属</td>
          <TD>管理操作</TD>
        </TR>
        <%Dim d_oneid,d_twoid,d_threeid
		If city_oneid=0 Then
        d_oneid=""
        Else
        d_oneid=" and city_oneid="&city_oneid&""
        End if
		If city_twoid=0 Then
        d_twoid=""
        Else
        d_twoid=" and city_twoid="&city_twoid&""
        End if
		If city_threeid=0 Then
        d_threeid=""
        Else
        d_threeid=" and city_threeid="&city_threeid&""
        End If
        selectm=trim(request("selectm"))
        selectkey=trim(request("selectkey"))
		set Rs=server.CreateObject("adodb.recordset")
		    sql= "select * from DNJY_bmcs Where bmname<>''"&d_oneid&""&d_twoid&""&d_threeid&"" 
		  	select case selectm
			case ""
            sql=sql
		    case "bmname"
			sql=sql&" and bmname like '%"&selectkey&"%'"
			case "url"
            sql=sql&" and url like '%"&selectkey&"%'"
            end Select
            sql=sql&" order by paixu "&DN_OrderType&"" 
			rs.open sql,conn,1,1
		  dim follows
		  if Rs.EOF and Rs.BOF Then
		  If city_oneid>0 Then dd=conn.Execute("Select city From DNJY_city Where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid&"")(0)
          response.write"<tr bgcolor=#FFFFFF><td colspan='6'><p align='center'><font color='red'>暂无"&dd&"便民服务！</font></td></tr></table><br>"
		  follows=0
		  Else
Dim Pagesize,Allrecord,Allpage,ThisPage,k,ss,selectkey,selectm,page
if request("page")="" then
  ThisPage=1		
else
  ThisPage=request("page")
end if
rs.Pagesize=15 '每页显示条数
Pagesize=rs.Pagesize
Allrecord=rs.Recordcount
Allpage=rs.Pagecount
if ThisPage<1 then                           
ThisPage=1
end if
rs.move (ThisPage-1)*Pagesize
k=1		  
		  do while not Rs.EOF
		  i=i+1%>
        <FORM name="form1" method="post" action="?action=edit&id=<%=int(Rs("id"))%>&city_oneid=<%=city_oneid%>&city_twoid=<%=city_twoid%>&city_threeid=<%=city_threeid%>">
          <TR bgcolor="#FFFFFF" align="center">
		  <TD><%=i%></TD>
		  <TD><INPUT name="bmname" type="text" id="bmname" value="<%=trim(Rs("bmname"))%>" size="12"></TD>
            <TD><INPUT name="link" type="text" id="link" value="<%=trim(Rs("link"))%>" size="20"></TD>
            <td><input name="paixu" type="text" id="paixu" value="<%=trim(rs("paixu"))%>" size="5"></td>
			<td><%If rs("city_oneid")=0 Then
            response.write"总站"
			Else
			response.write conn.Execute("Select city From DNJY_city Where id="&rs("city_oneid")&"")(0)&"/"&conn.Execute("Select city From DNJY_city Where id="&rs("city_oneid")&" and twoid="&rs("city_twoid")&"")(0)&"/"&conn.Execute("Select city From DNJY_city Where id="&rs("city_oneid")&" and twoid="&rs("city_twoid")&" and threeid="&rs("city_threeid")&"")(0)
			End if%></td>
            <TD><INPUT type="submit" name="Submit" value="修 改">
                &nbsp; <INPUT type="button" name="DEL" onClick="{if(confirm('确定要删除这个便民服务吗？\n此操作不可以恢复！')){location.href='?id=<%=Rs("id")%>&action=del&city_oneid=<%=city_oneid%>&city_twoid=<%=city_twoid%>&city_threeid=<%=city_threeid%>';}return false;}" value="删除" >            </TD>
          </TR>
        </FORM>
        <%Rs.MoveNext
	   if k>=Pagesize then exit Do
	   k=k+1
          loop
          follows=Rs.RecordCount
          end if
		  conn.close:Set conn=nothing%>
      </TABLE>
	<BR></TD>
  </TR>
</TABLE>
 <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#6699cc">
<tr> 
<td width="595" height="25" align="center">
  共有&nbsp;<font color="#CC5200"><%=Allrecord%></font>&nbsp;条记录 共 <font color="#CC5200"><%=Allpage%></font> 页 现在是第 <font color="#CC5200"><%=ThisPage%></font> 页
<%
if ThisPage<2 then     
response.write "<font color=""#808080"">首页</font>&nbsp;"
response.write "<font color=""#808080"">上一页</font>&nbsp;"     
else     
response.write "<a href=?page=1&ss="&ss&"&city_oneid="&city_oneid&"&city_twoid="&city_twoid&"&city_threeid="&city_threeid&"&selectkey="&selectkey&"&selectm="&selectm&">首页</a>&nbsp;"
response.write "<a href=?page="&ThisPage-1&"&ss="&ss&"&city_oneid="&city_oneid&"&city_twoid="&city_twoid&"&city_threeid="&city_threeid&"&selectkey="&selectkey&"&selectm="&selectm&">上一页</a>&nbsp;"     
end if
if Allpage-ThisPage<1 then     
response.write "<font color=""#808080"">下一页</font>&nbsp;"
response.write "<font color=""#808080"">尾页</font>&nbsp;"  
else     
response.write "<a href=?page="&(ThisPage+1)&"&ss="&ss&"&city_oneid="&city_oneid&"&city_twoid="&city_twoid&"&city_threeid="&city_threeid&"&selectkey="&selectkey&"&selectm="&selectm&">下一页</a>&nbsp;"   
response.write "<a href=?page="&Allpage&"&ss="&ss&"&city_oneid="&city_oneid&"&city_twoid="&city_twoid&"&city_threeid="&city_threeid&"&selectkey="&selectkey&"&selectm="&selectm&">尾页</a>&nbsp;"     
end if
response.write "每页"&Pagesize&"条记录"%></td>
</tr>
      </table> 
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#6699cc">
  <tr> 
    <td align="center" colspan="3">搜索查询</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <form name="form2" method="post" action="?page=<%=Page%>&city_oneid=<%=city_oneid%>&city_twoid=<%=city_twoid%>&city_threeid=<%=city_threeid%>">
      <td>
	  <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
          <tr bgcolor="#FFFFFF" align="center"> 
            <td width="35%"> <input name="selectkey" type="text" id="selectkey" onFocus="this.value=''" value="请输入关健字" size="30"> 
             <select name="selectm" id="select">                
                <OPTION VALUE="bmname">按网站名称</OPTION> 
                <OPTION VALUE="url">按网址</OPTION>
              </select> <input type="submit" name="Submit" value="查 询" ></td>
          </tr>
        </table></td>
    </form>
  </tr>
</table>	  
<BR>
<TABLE width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
  <TR> 
    <TD height="20" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">添 加 便 民 服 务</FONT></TD>
  </TR>
  <TR> 
    <TD bgcolor="#FFFFFF">
	<BR>
	<TABLE width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
        <TR align="center" bgcolor="#FFFFFF" height="20"> 
          <TD width="30">编号 </TD>
          <TD>网站名称</TD>
          <TD>网址链接</TD>
		  <td>排序</td>
          <TD>确定操作</TD>
        </TR>
        <FORM name="form1" method="post" action="?action=add&city_oneid=<%=city_oneid%>&city_twoid=<%=city_twoid%>&city_threeid=<%=city_threeid%>">
          <TR align="center" bgcolor="#FFFFFF"> 
            <TD><%=follows+1%></TD>
	        <TD><INPUT name="bmname" type="text" id="bmname" size="12"></TD>
	        <TD><INPUT name="link" type="text" id="link" size="20"></TD>
			<td><input name="paixu" type="text" id="paixu" size="5"></td>
	        <TD><INPUT type="submit" name="Submit3" value="添 加" onClick="if(bmname.value=='' || link.value==''){alert('不能为空');return false;}"></TD>
          </TR>
        </FORM>
      </TABLE>
	<BR></TD>
  </TR>
</TABLE>
</BODY> 
</HTML> 
