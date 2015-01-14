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
<%call checkmanage("06")
Call OpenConn
dim rsbigclass,rssmallclass,errmsg
set rsbigclass=server.createobject("adodb.recordset")
rsbigclass.open "select * from DNJY_bigClass order by BigClassID",conn,1,3
%>
<script language="javascript" type="text/javascript">
function checkbig()
{
  if (document.form1.bigclassname.value=="")
  {
    alert("大类名称不能为空！");
    document.form1.bigclassname.focus();
    return false;
  }
}
function checksmall()
{
  if (document.form2.bigclassname.value=="")
  {
    alert("请先添加大类名称！");
	document.form1.bigclassname.focus();
	return false;
  }

  if (document.form2.smallclassname.value=="")
  {
    alert("小类名称不能为空！");
	document.form2.smallclassname.focus();
	return false;
  }
}
function confirmdelbig()
{
   if(confirm("确定要删除此大类吗？删除此大类同时将删除所包含的小类和该类下的所有文章，并且不能恢复！"))
     return true;
   else
     return false;
	 
}

function confirmdelsmall()
{
   if(confirm("确定要删除此小类吗？删除此小类同时将删除该类下的所有文章，并且不能恢复！"))
     return true;
   else
     return false;
	 
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
<BODY background="img/back.gif">
<script language=javascript src=../Include/mouse_on_title.js></script>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr class=backs><td align="center" class=td height=18 colspan="8"><font color="#FFFFFF">文章分类管理</font></td>
  </tr>
  <tr> 
    <td align="center" valign="top"><br>
      <%if session("aleave")="check" then%><%else%>
	  <a href="classaddbig.asp"><strong><font color="#ff0000"><u>添加文章一级分类</u></font></strong></a><br><%end if%>
      <br> 
      <table width="98%" border="0" cellpadding="0" cellspacing="1" bgcolor="#000000">
        <tr> 
          <td width="40%" height="30" align="center" bgcolor="#999999"><strong>栏目名称</strong></td>		  
		  <td width="10%" height="30" align="center" bgcolor="#999999"><strong>分类ID<br>（classid）</strong></td>
		  <td width="10%" height="30" align="center" bgcolor="#999999"><strong>排序</strong></td>
		  <td width="10%" height="30" align="center" bgcolor="#999999"><strong>是否允许会员投稿</strong></td>
          <td height="40" align="center" bgcolor="#999999"><strong>操作选项</strong></td>
        </tr>
        <%
	do while not rsbigclass.eof
%>
        <tr bgcolor="#e3e3e3" class="tdbg"> 
          <td width="233" height="22" bgcolor="#c0c0c0"><img src="../img/zj.gif" width="15" height="15"><%=rsbigclass("bigclassname")%> </td>
		  <td width="233" height="22" bgcolor="#c0c0c0"><font color="#FF0000"><%=rsbigclass("id")%></font></td>
          <td width="233" height="22" bgcolor="#c0c0c0"><font color="#0000FF"><%=rsbigclass("bigclassid")%></font></td>
          <td width="233" height="22" bgcolor="#c0c0c0"><%If rsbigclass("userdraft")=1 then Response.Write("<font color='#0000FF'>√</font>") else Response.Write("<font color=red>ㄨ</font>") end if%></td>
          <td  bgcolor="#c0c0c0" style="padding-right:10"> 
	<%if session("aleave")="check" then
		  else
		  if rsbigclass("urlok")=false then
		  %>
		  <a href="classaddsmall.asp?bigclassname=<%=rsbigclass("bigclassname")%>"><font color="#ff0000">添加二级分类</font></a> 				  		  <%end if
	 end if%>
            | <%if session("aleave")="check" then%>修改<%else%><a href="classmodifybig.asp?bigclassid=<%=rsbigclass("bigclassid")%>">修改</a><%end if%> 
            | <%if session("aleave")="check" then%>删除<%else%><a href="classdelbig.asp?bigclassname=<%=rsbigclass("bigclassname")%>" onclick="return confirmdelbig();">删除</a><%end if%></td>
        </tr>
        <%
	  set rssmallclass=server.createobject("adodb.recordset")
	  rssmallclass.open "select * from DNJY_SmallClass where bigclassname='" & rsbigclass("bigclassname") & "'  order by sortingid",conn,1,1
	  if not(rssmallclass.bof and rssmallclass.eof) then
		do while not rssmallclass.eof
	%>
        <tr bgcolor="#eaeaea" class="tdbg"> 
          <td width="233" height="22">&nbsp;&nbsp;<img src="../img/js.gif" width="15" height="15"><%=rssmallclass("smallclassname")%></td>          
          <td width="233" height="22" bgcolor="#c0c0c0"><font color="#1DA2DB">&nbsp;&nbsp;<%=rssmallclass("SmallClassID")%></font></td>
		  <td width="233" height="22" bgcolor="#c0c0c0"><font color="#1DA2DB">&nbsp;&nbsp;<%=rssmallclass("sortingid")%></font></td>
		  <td width="233" height="22" bgcolor="#c0c0c0"><%If rssmallclass("userdraft")=1 then Response.Write("<font color='#0000FF'>√</font>") else Response.Write("<font color=red>ㄨ</font>") end if%></td>
          <td style="padding-right:10"><%if session("aleave")="check" then%>修改<%else%><a href="classmodifysmall.asp?smallclassid=<%=rssmallclass("smallclassid")%>">修改</a><%end if%> 
            | <%if session("aleave")="check" then%>删除<%else%><a href="classdelsmall.asp?smallclassid=<%=rssmallclass("smallclassid")%>&smallclassname=<%=rssmallclass("smallclassname")%>" onclick="return confirmdelsmall();">删除</a><%end if%></td>
        </tr>
        <%
			rssmallclass.movenext
		loop
	  end if
	  rssmallclass.close
	  set rssmallclass=nothing	
	  rsbigclass.movenext
	loop
%>
      </table>
      <br>
    </td>
  </tr>
</table>
<%
rsbigclass.close
set rsbigclass=Nothing
Call CloseDB(conn)%>


                                                              
                                                              
                                                              