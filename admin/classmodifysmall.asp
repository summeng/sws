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
<%
dim smallclassid,action,bigclassname,smallclassname,oldsmallclassname,founderr,errmsg,sortingid
smallclassid=trim(request("smallclassid"))
sortingid=trim(request("sortingid"))
action=trim(request("action"))
bigclassname=trim(request.form("bigclassname"))
smallclassname=trim(request.form("smallclassname"))
oldsmallclassname=trim(request.form("oldsmallclassname"))
if smallclassid="" then
	response.redirect("classmanage.asp")
end If
Call OpenConn
set rs=server.createobject("adodb.recordset")
rs.open "select * from DNJY_SmallClass where smallclassid="&smallclassid&"",conn,1,3
if rs.bof or rs.eof then
	founderr=true
	errmsg=errmsg & "<br><li>此文章小类不存在！</li>"
else
	if action="modify" then
		if smallclassname="" then
			founderr=true
			errmsg=errmsg & "<br><li>文章小类名不能为空！</li>"
		end if
		if founderr<>true then
			rs("smallclassname")=smallclassname
            rs("sortingid")=sortingid
			rs("userdraft")=trim(request("userdraft"))
     		rs.update
			rs.close
    	 	set rs=nothing
			if smallclassname<>oldsmallclassname then
				conn.execute "update DNJY_News set smallclassname='" & smallclassname & "' where bigclassname='" & bigclassname & "' and smallclassname='" & oldsmallclassname & "'"
	     	end if	
			'call closeconn()
    	 	response.redirect "classmanage.asp"
		end if
	end if
	if founderr=true then
		call writeerrmsg()
	else
%>
<script language="javascript" type="text/javascript">
function checksmall()
{
  if (document.form1.smallclassname.value=="")
  {
    alert("小类名称不能为空！");
	document.form1.smallclassname.focus();
	return false;
  }
  if (document.form1.sortingid.value=="")
  {
    alert("小类排序不能为空！");
	document.form1.sortingid.focus();
	return false;
  }
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
<BODY>
<script language=javascript src=../Include/mouse_on_title.js></script>
<table width="80%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> <br>
      <a href="classmanage.asp"><strong>新 闻 类 别 设 置</strong></a> <br> <form name="form1" method="post" action="classmodifysmall.asp" onsubmit="return checksmall()">
        <p>&nbsp;</p>
        <table width="350" border="0" align="center" cellpadding="0" cellspacing="2" class="border">
          <tr bgcolor="#999999" class="title"> 
            <td height="25" colspan="2" align="center"><strong>修改小类名称</strong></td>
          </tr>
          <tr class="tdbg"> 
            <td width="126" height="22" bgcolor="#c0c0c0" align="right"><strong>所属大类：</strong></td>
            <td width="204" bgcolor="#e3e3e3"><%=rs("bigclassname")%> <input name="smallclassid" type="hidden" id="smallclassid" value="<%=rs("smallclassid")%>"> 
              <input name="oldsmallclassname" type="hidden" id="oldsmallclassname" value="<%=rs("smallclassname")%>"> 
              <input name="bigclassname" type="hidden" id="bigclassname" value="<%=rs("bigclassname")%>"></td>
          </tr>
          <tr class="tdbg"> 
            <td height="22" bgcolor="#c0c0c0" align="right"><strong>小类名称：</strong></td>
            <td bgcolor="#e3e3e3"> <input name="smallclassname" type="text" id="smallclassname" value="<%=rs("smallclassname")%>" size="20" maxlength="30"></td>
          </tr>
          <tr class="tdbg"> 
            <td height="22" bgcolor="#c0c0c0" align="right"><strong>小类排序：</strong></td>
            <td bgcolor="#e3e3e3"> <input name="sortingid" type="text" id="sortingid" value="<%=rs("sortingid")%>" size="20" maxlength="30" onKeyUp="if(isNaN(value)){alert('只允许输入数字');value='';}"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))"></td>
          </tr>
		  <tr bgcolor="#e3e3e3" class="tdbg"> 
            <td width="126" height="22" bgcolor="#c0c0c0"> <div align="right"><strong>是否允许会员投稿：</strong></div></td>
            <td width="218" bgcolor="#e3e3e3"><label><input name="userdraft" type="radio" value="1"<%If rs("userdraft")=1 then Response.Write(" checked") end if%>>是  <input name="userdraft" type="radio" value="0"<%If rs("userdraft")=0 then Response.Write(" checked") end if%>>
否</label></td>
          </tr>
          <tr class="tdbg"> 
            <td height="22" align="center" bgcolor="#c0c0c0">&nbsp;</td>
            <td align="center" bgcolor="#e3e3e3"> <div align="left"> 
                <input name="action" type="hidden" id="action4" value="modify">
                <input name="save" type="submit" id="save" value=" 修 改 ">
              </div></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
<%
end if
end if
Call CloseRs(rs)
Call CloseDB(conn)%>


                                                              
                                                              
                                                              