<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("06")%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%Call OpenConn
If request("yz")="yes" Or request("yz")="no" Then
Dim str2,i
Server.ScriptTimeout=50000
id=trim(request("DelID"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('没有选择记录！');" &"window.location='admin_article.asp'" & "</script>"
response.end
end if
str2=split(id,",")
for i=0 to ubound(str2)
set rs=server.createobject("adodb.recordset")
sql="select ifshow from [DNJY_news] where id="&trim(str2(i))
rs.open sql,conn,1,3
If request("yz")="yes" Then
rs("ifshow")=1
Else
rs("ifshow")=0
End if
rs.update
Next
Call CloseRs(rs)
End If

dim rs5,rsc,count,page,pages,j,rsfl1,bigid,SmallCID,SmallClassName,rsfl2,sql1,isNaN
set rs5=server.createobject("adodb.recordset")
sql = "select * from DNJY_SmallClass order by sortingid"
rs5.open sql,conn,1,1
%>
<script language = "JavaScript">
var onecount;
subcat = new Array();
        <%
        count = 0
        do while not rs5.eof 
        %>
subcat[<%=count%>] = new Array("<%= trim(rs5("SmallClassName"))%>","<%= trim(rs5("bigid"))%>","<%= trim(rs5("SmallClassID"))%>");
        <%
        count = count + 1
        rs5.movenext
        loop
        rs5.close
        %>
onecount=<%=count%>;

function changelocation(locationid)
    {
    document.addNEWS.SmallClassID.length = 1; 
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.addNEWS.SmallClassID.options[document.addNEWS.SmallClassID.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
    }    


</script>
<title>管理文章</title>
<link rel="stylesheet" href="../css/manage.css" type="text/css">
<script language=javascript src=../Include/mouse_on_title.js></script>
<script language=javascript>
<!--
function left_menu(meval)
{
  var left_n=eval(meval);
  if (left_n.style.display=="none")
  { eval(meval+".style.display='';"); }
  else
  { eval(meval+".style.display='none';"); }
}
function DoEmpty(params)
{
if (confirm("真的要删除这条记录吗？删除后将无法恢复！"))
window.location = params ;
}
function test()
{
  if(!confirm('是否确定进行批量操作？操作后不能恢复！')) return false;
}
-->
</script>
<script language="JavaScript">
<!--
function CheckAll(form)
{
for (var i=0;i<form.elements.length;i++)
{
var e = form.elements[i];
if (e.name != 'chkall')
e.checked = form.chkall.checked;
}
}

//-->
</script>
<SCRIPT language=JavaScript>
function showoperatealert(id)
{
		if (id==1)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="admin_article.asp?yz=yes";
			thisForm.submit();
		}
		}
		if (id==2)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="admin_article.asp?yz=no";
			thisForm.submit();
		}
		}
		if (id==3)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="admin_article_del.asp?action=delarticle";
			thisForm.submit();
		}
		}
}

//-->
    </SCRIPT>
</head>
<BODY background="img/back.gif"><br>

<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#000000">
<tr class=backs><td align="center" class=td height=18 colspan="22"><font color="#FFFFFF">文章管理</font></td>
  </tr>


          <tr bgcolor="#FFFFFF"> 
 <td colspan="6" width="70%">	  <form name="addNEWS" method="post" action="admin_article.asp?no=modi" >
	        <%set rsc=server.createobject("ADODB.Recordset")
        sql = "select * from DNJY_bigClass where urlok=0 order by BigClassID"
        rsc.open sql,conn,1,1
		if rsc.eof and rsc.bof then
		response.write "请先添加栏目。"
        response.end
		else
		%>按分类查询:
        <select name="id" onChange="changelocation(document.addNEWS.id.options[document.addNEWS.id.selectedIndex].value)" size="1">
          
          <option selected value="<%=trim(rsc("id"))%>"><%=trim(rsc("BigClassName"))%></option>
          <%
			dim selclass
		    selclass=rsc("id")
        	rsc.movenext
		    do while not rsc.eof
			%>
          <option value="<%=trim(rsc("id"))%>"><%=trim(rsc("BigClassName"))%></option>
          <%
		        rsc.movenext
    	    loop
		end if
        rsc.close
			%>
        </select> <select name="SmallClassID">
          <option value="" selected>不指定小类</option>
          <%
			sql="select * from DNJY_SmallClass where bigid="&selclass&" order by sortingid"
			rsc.open sql,conn,1,1
			if not(rsc.eof and rsc.bof) then
			%>
          <option value="<%=rsc("SmallClassID")%>"><%=rsc("SmallClassName")%></option>
          <% rsc.movenext
				do while not rsc.eof%>
          <option value="<%=rsc("SmallClassID")%>"><%=rsc("SmallClassName")%></option>
          <%
			    rsc.movenext
				loop
			end if
	        rsc.close
			%>         
        </select>
        <input type="submit" name="Submit" value="提交" class="input">		
		
		</form>
 </td>
 

 <td width="30%" colspan="10">按类型查询: 		<select name="selectm" onChange="var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location=jmpURL;} else {this.selectedIndex=0 ;}" target=_blank>
			    <OPTION value="admin_article.asp?recommend=""">类型</OPTION>
				<OPTION value="admin_article.asp?recommend=0">全部</OPTION> 
		        <OPTION value="admin_article.asp?recommend=1">推荐</OPTION>
                <OPTION value="admin_article.asp?recommend=2">置顶</OPTION>
                <OPTION value="admin_article.asp?recommend=3">待审</OPTION>                
                           
     </select> 
            </td>       
          </tr>


 <FORM name="thisForm" action="admin_article_del.asp?action=delarticle" method="POST">
  <tr> 
    <td width="3%" height="25" align="center" bgcolor="#cccccc">选择</td>
    <td width="3%" align="center" bgcolor="#cccccc">评论</td>
    <td width="30%" align="center" bgcolor="#cccccc">文章标题</td>
    <td width="8%" align="center" bgcolor="#cccccc">城市</td>
    <td width="8%" align="center" bgcolor="#cccccc">发布者</td>
    <td width="15%" align="center" bgcolor="#cccccc">发布日期</td>
    <td width="5%" align="center" bgcolor="#cccccc">推荐</td>
    <td width="5%" align="center" bgcolor="#cccccc">置顶</td>
    <td width="5%" align="center" bgcolor="#cccccc">评论</td>
    <td width="13%" align="center" bgcolor="#cccccc">操作</td>
  </tr>
  <%
page=clng(request("page"))
dim selectm,selectkey,selectid,ProdId,comment,id
selectkey=trim(request.form("selectkey"))
selectm=trim(request.form("selectm"))
if selectkey="" then
selectkey=request.QueryString("selectkey")
end if
if selectm="" then
selectm=request.QueryString("selectm")
end if

set rs=server.createobject("adodb.recordset")
if Request.QueryString("recommend")=""   Then
sql="select * from DNJY_News  "
if (request("no")="modi" Or request("fy")="ok") And (request("id")<>"") Then
sql=sql&"where bigcid="&trim(request("id"))&""
If request("SmallClassID")<>"" Then
sql=sql&" and SmallCID="&trim(request("SmallClassID"))&""
End If
End if
else
sql="select * from DNJY_News  "	
if Request.QueryString("recommend")=1 then
sql=sql&"where ifshow=1 and tuijian=1 "
Else
if Request.QueryString("recommend")=2 then
sql=sql&"where ifshow=1 and newstop=1 "
Else
if Request.QueryString("recommend")=3 then
sql=sql&"where  ifshow=0 "
Else
sql=sql
End If
End If
End If
End If


			select case selectm
			case ""
            sql=sql
		    case "name"
			sql=sql&" where title like '%"&selectkey&"%'"
			case "ProdId"
   If  Not isNumeric(selectkey) Then  
   response.write "编号应为数字！"     
   response.End
   else
			sql=sql&" where   Id =cint("&selectkey&")"
   End if
			case "content"
			sql=sql&"where  content like '%"&selectkey&"%'"
			
		  end select


sql=sql&" order by id "&DN_OrderType&""
rs.open sql,conn,1,1
if rs.eof and rs.bof then
%>
<tr bgcolor="#FFFFFF"><td colspan="22" align="center">暂时没有相关记录</td>
</tr>
<%response.end
Else
rs.pagesize=15
if page=0 then page=1 
pages=rs.pagecount
if page > pages then page=pages
rs.absolutepage=page  
    
for j=1 to rs.pagesize 
Dim BigClassName
set rsfl1=server.createobject("adodb.recordset")
sql="select * from DNJY_bigClass where  ID="&rs("bigcid")&" "
rsfl1.open sql,conn,1,1
if rsfl1.eof and rsfl1.bof Then
response.write "没有此分类"
else
BigClassName=rsfl1("BigClassName")
SmallCID=CInt(rs("SmallCID"))
If SmallCID<>0 Then
set rsfl2=server.createobject("adodb.recordset")
sql="select * from DNJY_SmallClass where  SmallClassID="&rs("SmallCID")&" "
rsfl2.open sql,conn,1,1
if Not rsfl2.eof Then SmallClassName=rsfl2("SmallClassName")
rsfl2.close
set rsfl2=nothing
Else
SmallClassName=""
End If
End If
rsfl1.close
set rsfl1=Nothing


	  dim rs1
		set rs1=server.createobject("adodb.recordset") 
		sql1="select * from DNJY_news_pinglun where id="&rs("id")&" order by pinglunid "&DN_OrderType&""
		rs1.open sql1,conn,1,1
%>
  <tr bgcolor="#ffffff"> 
    <td height="22" align="center"><input name="DelID" type="checkbox" id="DelID" value="<%=rs("id")%>"></td>
	<td align="center"><img onClick="javascript:left_menu('left_<%=rs("id")%>');" src="img/dedeexplode.gif" alt="展开评论" width="11" height="11"></td>
    <td>[<b><%=BigClassName%><%If SmallClassName<>"" then%>--<%=SmallClassName%><%End if%></b>] <a class=<%=rs("oColor")%>
news_show.asp?catid=0&bigcid="&Cstr(rs("bigcid"))&"&Id="&Cstr(rs("Id"))&"&"&L_id&" 
 href="../news_show.asp?id=<%=rs("id")%>&classid=<%=rs("bigcid")%>&bigcid=<%=rs("bigcid")%>" target="_blank" title="<%=rs("title")%>"><span class="<%=rs("oStyle")%>"><%=left(rs("title"),20)%></span></a><% if rs("img")=True then response.write "<img src='img/pic.gif' border=0 alt='图片文章'>" end if %></td>
     <TD align="center"><A title=<%=rs("city_one")%>/<%=rs("city_two")%>/<%=rs("city_three")%> href="#"><%=rs("city_one")%></A></TD>
    <td align="center"><%=rs("newsuser")%></td>
    <td align="center"><%=rs("infotime")%></td>
    <td align="center"><%if rs("tuijian")=1 then%><a href=admin_article_audit.asp?action=tuijianno&id=<%=rs("id")%>><font color=red>是</font></a><%else%><a href=admin_article_audit.asp?action=tuijianyes&id=<%=rs("id")%>>否</a><%end if%></td>
    <td align="center">
	<%if rs("newstop")=1 then%><a href=admin_article_audit.asp?action=headno&id=<%=rs("id")%>><font color=red>是</font></a><%else%><a href=admin_article_audit.asp?action=headyes&id=<%=rs("id")%>>否</a><%end if%></td>
    <td align="center"><%=rs1.recordcount%></td>

    <td align="center"><%if rs("ifshow")=1 then%><a href="admin_article_audit.asp?action=ifshowno&id=<%=rs("id")%>">已审</a><%else%><a href="admin_article_audit.asp?action=ifshowyes&id=<%=rs("id")%>"><font color=red>待审</font></a><%end if%> <a href="admin_article_modi.asp?id=<%=rs("id")%>">修改</a> 
	<%if session("aleave")="check" then%>删除<%else%><a href="javascript:DoEmpty('admin_article_del.asp?DelID=<%=rs("id")%>');">删除</a><%end if%></td>
  </tr>
  <tr bgcolor="#ffffff" id="left_<%=rs("id")%>" style="display:none;">
    <td height="22" colspan="10" align="right">
	<table width="95%" border="0" cellpadding="2" cellspacing="1" bgcolor="#999999">
      <tr>
        <td width="4%" height="25" align="center" bgcolor="#cccccc">id</td>
        <td width="8%" align="center" bgcolor="#cccccc">用户名称</td>
        <td width="55%" align="center" bgcolor="#cccccc">评论内容</td>
        <td width="15%" align="center" bgcolor="#cccccc">评论时间</td>
        <td width="8%" align="center" bgcolor="#cccccc">审核</td>
        <td width="8%" align="center" bgcolor="#cccccc">操作</td>
      </tr>
	  <%

		if rs1.eof and rs1.bof then
		
		%>
		 <tr>
        <td height="25" colspan="10" align="center" bgcolor="#f8f8f8">暂时还没有评论</td>
        </tr>
		<%else
		do while not rs1.eof 
	  %>
      <tr>
        <td height="25" align="center" bgcolor="#f8f8f8"><%=rs1("pinglunid")%></td>
        <td align="center" bgcolor="#f8f8f8"><%=rs1("pinglunname")%></td>
        <td bgcolor="#f8f8f8"><%=rs1("pingluncontent")%></td>
        <td align="center" bgcolor="#f8f8f8"><%=rs1("pinglundate")%></td>
       <td align="center" bgcolor="#f8f8f8">
	<%if rs1("sh")=1 then%><a href=admin_pinglunsh.asp?action=shno&id=<%=rs1("pinglunid")%>><font color=red>是</font></a><%else%><a href=admin_pinglunsh.asp?action=sh&id=<%=rs1("pinglunid")%>>否</a><%end if%></td>
        <td align="center" bgcolor="#f8f8f8"><%if session("aleave")="check" then%>删除<%else%><a href="javascript:DoEmpty('admin_pinglundel.asp?id=<%=rs1("pinglunid")%>');">删除</a><%end if%></td>
      </tr>
     
	  <%rs1.movenext
		loop
		rs1.close
		set rs1=nothing
		end if%>
    </table>	</td>
  </tr>
  <%
rs.movenext
if rs.eof then exit For
next                                                       
%>
</table>
  <p align="center">
    <input type="checkbox" name="chkall" onClick="CheckAll(this.form)">
    全部选中
<input onclick=javascript:showoperatealert(1) type="submit" value="批量审核通过" name="B1">
<input onclick=javascript:showoperatealert(2) type="submit" value="批量取消通过" name="B2">	  
<input name='submit1' type='submit' value=' 批量删除 ' onClick="return test();"  style="color: #FF0000">	
  </p>
</form>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr bgcolor="#ffffff"> 
<form method=post action="admin_article.asp?recommend=<%=Request.QueryString("recommend")%>">  
      <td height="30" align="center"> 
    <%if page<2 then      
    response.write "首页 上一页&nbsp;"
  else
    response.write "<a href=admin_article.asp?page=1&recommend="&Request.QueryString("recommend")&"&id="&trim(request("id"))&"&SmallClassID="&trim(request("SmallClassID"))&"&fy=ok&selectm="&selectm&"&selectkey="&selectkey&">首页</a>&nbsp;"
    response.write "<a href=admin_article.asp?page=" & page-1 & "&recommend="&Request.QueryString("recommend")&"&id="&trim(request("id"))&"&SmallClassID="&trim(request("SmallClassID"))&"&fy=ok&selectm="&selectm&"&selectkey="&selectkey&">上一页</a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "下一页 尾页"
  else
    response.write "<a href=admin_article.asp?page=" & (page+1) & "&recommend="&Request.QueryString("recommend")&"&id="&trim(request("id"))&"&SmallClassID="&trim(request("SmallClassID"))&"&fy=ok&selectm="&selectm&"&selectkey="&selectkey&">"
    response.write "下一页</a> <a href=admin_article.asp?page="&rs.pagecount&"&recommend="&Request.QueryString("recommend")&"&id="&trim(request("id"))&"&SmallClassID="&trim(request("SmallClassID"))&"&fy=ok&selectm="&selectm&"&selectkey="&selectkey&">尾页</a>"
  end if
   response.write "&nbsp;页次：<strong><font color=red>"&page&"</font>/"&rs.pagecount&"</strong>页 "
    response.write "&nbsp;共<b><font color='#ff0000'>"&rs.recordcount&"</font></b>条记录 <b>"&rs.pagesize&"</b>条记录/页"
   response.write " 转到：<input type='text' name='page' size=4 maxlength=10 class=input value="&page&">"
   response.write " <input class=input type='submit'  value=' goto '  name='cndok'></span></p>"     
%>
      </td></form>
  </tr>
</table>
<%end if
Call CloseRs(rs)
Call CloseDB(conn)%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#6699cc">
  <tr> 
    <td align="center">搜索查询</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <form name="form2" method="post" action="admin_article.asp">
      <td>
	  <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
          <tr bgcolor="#FFFFFF" align="center"> 
            <td > <input name="selectkey" type="text" id="selectkey" onFocus="this.value=''" value="请输入关健字" size="30"> 
             <select name="selectm" id="select">			   
                <OPTION VALUE="name">按标题</OPTION>                
                <OPTION VALUE="content">按内容</OPTION>
				<OPTION VALUE="ProdId">按编号</OPTION> 
              </select><input type="submit" name="Submit" value="查 询" ></td>
          </tr>
        </table></td>
    </form>
  </tr>
</table>

</body>
</html>
