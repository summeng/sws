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
dim rs2,bigclassid,action,newbigclassname,oldbigclassname,founderr,errmsg,newbigclassid,hide,url,admin
bigclassid=trim(request("bigclassid"))
newbigclassid=trim(request("newbigclassid"))
action=trim(request("action"))
newbigclassname=trim(request("newbigclassname"))
oldbigclassname=trim(request("oldbigclassname"))
hide=request("hide")
url=trim(request("url"))
if bigclassid="" then
  response.redirect("classmanage.asp")
end If
Call OpenConn
set rs=server.createobject("adodb.recordset")
rs.open "select * from DNJY_bigClass where bigclassid=" & clng(bigclassid),conn,1,3
if rs.bof and rs.eof then
	founderr=true
	errmsg=errmsg & "<br><li>此新闻大类不存在！</li>"
else
	if action="modify" then
	if request("newbigclassid")<>request("bigclassid") then
		dim classid
		set rs2 = server.CreateObject ("adodb.recordset")
		sql="select bigclassid from DNJY_bigClass where bigclassid="+cstr(request("newbigclassid"))
 		rs2.open sql,conn,1,1
		if not rs2.eof and not rs2.bof then
		rs2.close
		set rs2=nothing%>
		<script language=javascript>  
			alert( "错误：在类排序有重复"  );
			location.href = "javascript:history.back()"  
		</script>
		<%
		response.End()
		end if
		end if
		if newbigclassname="" then
			founderr=true
			errmsg=errmsg & "<br><li>新闻大类名不能为空！</li>"
		end if
		if founderr<>true then
			rs("bigclassname")=newbigclassname
			rs("bigclassid")=newbigclassid
			rs("admin")=admin
			if hide="yes" then
			rs("hide")=true
			else
			rs("hide")=false
			end if
			rs("url")=url
			if url<>"" then
			rs("urlok")=true
			else
			rs("urlok")=false
			end If
			rs("userdraft")=trim(request("userdraft"))
    	 	rs.update
			rs.close
	     	set rs=nothing
			if newbigclassname<>oldbigclassname then
				conn.execute "update DNJY_SmallClass set bigclassname='" & newbigclassname & "' where bigclassname='" & oldbigclassname & "'"
				conn.execute "update DNJY_News set bigclassname='" & newbigclassname & "' where bigclassname='" & oldbigclassname & "'"
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
function checkbig()
{
  if (document.form1.newbigclassname.value=="")
  {
    alert("大类名称不能为空！");
    document.form1.newbigclassname.focus();
    return false;
  }
  if (document.form1.newbigclassid.value=="")
  {
    alert("大类排序不能为空！");
    document.form1.newbigclassid.focus();
    return false;
  }
}
</script>
<link href="../css/style.css" rel="stylesheet" type="text/css">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"><br> <a href="classmanage.asp"><strong>类 别 设 置</strong></a> <br> 
    <br> <form name="form1" method="post" action="classmodifybig.asp" onsubmit="return checkbig()">
        <table width="80%" border="0" align="center" cellpadding="0" cellspacing="2" class="border">
          <tr class="title"> 
            <td height="25" colspan="2" align="center" bgcolor="#999999"><strong>修改大类名称</strong></td>
          </tr>
          <tr class="tdbg"> 
            <td width="126" bgcolor="#c0c0c0" align="right"><strong>大类名称：</strong></td>
            <td bgcolor="#e3e3e3"> <input name="newbigclassname" type="text" id="newbigclassname" value="<%=rs("bigclassname")%>" size="20" maxlength="30"></td>
          </tr>
          <tr class="tdbg"> 
            <td width="126" bgcolor="#c0c0c0" align="right"><strong>大类排序：</strong></td>
            <td width="204" bgcolor="#e3e3e3"> 
			<input name="newbigclassid" type="text" id="newbigclassid" value="<%=rs("bigclassid")%>" size="20" maxlength="20" onKeyUp="if(isNaN(value)){alert('只允许输入数字');value='';}"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))">
			<input name="bigclassid" type="hidden" id="bigclassid" value="<%=rs("bigclassid")%>" size="20" maxlength="20" > 
              <input name="oldbigclassname" type="hidden" id="oldbigclassname" value="<%=rs("bigclassname")%>"></td>
          </tr>
		  <tr bgcolor="#e3e3e3" class="tdbg"> 
            <td height="22" bgcolor="#c0c0c0" align="right"> <strong>外链接地址：</strong></td>
            <td bgcolor="#e3e3e3"> <input name="url" type="text" id="url" value="<%=rs("url")%>" size="20" maxlength="50"> 
            (如果填写地址只作外部链接,不作内部新闻资讯栏目! )  </td>
          </tr>
		   <tr bgcolor="#e3e3e3" class="tdbg"> 
            <td height="22" bgcolor="#c0c0c0" align="right"><strong>首页显示：</strong></td>
            <td width="218" bgcolor="#e3e3e3"><label>
              <input name="hide" type="checkbox" id="hide" value="yes" <%if rs("hide")=DN_True then%>checked<%end if%> />
            （指首页版块）</label></td>
          </tr>
		  <tr bgcolor="#e3e3e3" class="tdbg"> 
            <td width="126" height="22" bgcolor="#c0c0c0"> <div align="right"><strong>是否允许会员投稿：</strong></div></td>
            <td width="218" bgcolor="#e3e3e3"><label><input name="userdraft" type="radio" value="1"<%If rs("userdraft")=1 then Response.Write(" checked") end if%>>是  <input name="userdraft" type="radio" value="0"<%If rs("userdraft")=0 then Response.Write(" checked") end if%>>
否</label></td>
          </tr>
          <tr class="tdbg"> 
            <td align="center" bgcolor="#c0c0c0">&nbsp;</td>
            <td align="center" bgcolor="#e3e3e3"> <div align="left"> 
                <input name="action" type="hidden" id="action" value="modify">
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


                                                              
                                                              
                                                              