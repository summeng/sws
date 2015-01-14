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
<%Call OpenConn
dim action,bigclassname,smallclassname,founderr,errmsg,bigid,rsbig,sortingid
action=trim(request("action"))
bigclassname=trim(request("bigclassname"))
smallclassname=trim(request("smallclassname"))
sortingid=trim(request("sortingid"))
if action="add" then
	if bigclassname="" then
		founderr=true
		errmsg=errmsg & "<br><li>文章大类名不能为空！</li>"
	end if
	if smallclassname="" then
		founderr=true
		errmsg=errmsg & "<br><li>文章小类名不能为空！</li>"
	end if
	if founderr<>true then
		set rs=server.createobject("adodb.recordset")
		rs.open "select * from DNJY_SmallClass where bigclassname='" & bigclassname & "' and smallclassname='" & smallclassname & "'",conn,1,3
		if not rs.eof then
			founderr=true
			errmsg=errmsg & "<br><li>“" & bigclassname & "”中已经存在文章小类“" & smallclassname & "”！</li>"
		Else
'-------取大类ID号
		set rsbig=server.createobject("adodb.recordset")
		rsbig.open "select * from DNJY_bigClass where BigClassName='"&bigclassname&"'",conn,1,3
		bigid=rsbig("ID")
	    rsbig.close
	    set rsbig=Nothing
'-------取大类ID号END	    
     		rs.addnew
			rs("bigclassname")=bigclassname
    	 	rs("smallclassname")=smallclassname
    	 	rs("bigid")=bigid
            rs("sortingid")=sortingid
			rs("userdraft")=trim(request("userdraft"))
     		rs.update
	     	rs.close
    	 	set rs=nothing
     		'call closeconn()
			response.redirect "classmanage.asp"  
		end if
	end if
end if
if founderr=true then
	response.Write ""&errmsg&""
else
%>
<script language="javascript" type="text/javascript">
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
  if (document.form2.sortingid.value=="")
  {
    alert("小类排序不能为空！");
	document.form2.sortingid.focus();
	return false;
  }
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
<BODY>
<script language=javascript src=../Include/mouse_on_title.js></script>
<%
'-------------------------------------
'天天网络商城
'天天DV网版权所有 http://www.ttdv.cn
'技术支持论坛 http://bbs.ip126.com/
'QQ:530051328 mail:bobai@126.com

'-------------------------------------%>
<table width="80%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> <br>
      <a href="classmanage.asp"><strong>新 闻 类 别 设 置</strong></a> <br> <form name="form2" method="post" action="classaddsmall.asp?bigid=bigid" onsubmit="return checksmall()">
        <table width="350" border="0" align="center" cellpadding="0" cellspacing="2" class="border">
          <tr bgcolor="#999999" class="title"> 
            <td height="25" colspan="2" align="center"><strong>添加小类</strong></td>
          </tr>
          <tr class="tdbg"> 
            <td width="126" height="22" bgcolor="#c0c0c0" align="right"><strong>所属大类：</strong></td>
            <td width="218" bgcolor="#e3e3e3"> <select name="bigclassname">
                <%
	dim rsbigclass
	set rsbigclass=server.createobject("adodb.recordset")
	rsbigclass.open "select * from DNJY_bigClass",conn,1,1
	if rsbigclass.bof and rsbigclass.bof then
		response.write "<option>请先添加文章大类</option>"
	else
		do while not rsbigclass.eof
			if rsbigclass("bigclassname")=bigclassname then
				response.write "<option value='"& rsbigclass("bigclassname") & "' selected>" & rsbigclass("bigclassname") & "</option>"
			else
				response.write "<option value='"& rsbigclass("bigclassname") & "'>" & rsbigclass("bigclassname") & "</option>"
			end if
			rsbigclass.movenext
		loop
	end if

		rsbigclass.close
	set rsbigclass=nothing%>
              </select></td>
          </tr>
          <tr class="tdbg"> 
            <td width="126" height="22" bgcolor="#c0c0c0" align="right"><strong>小类名称：</strong></td>
            <td bgcolor="#e3e3e3"> <input name="smallclassname" type="text" size="20" maxlength="30"></td>
          </tr>
          <tr class="tdbg"> 
            <td width="126" height="22" bgcolor="#c0c0c0" align="right"><strong>小类排序：</strong></td>
            <td bgcolor="#e3e3e3"> <input name="sortingid" type="text" size="10" maxlength="10" onKeyUp="if(isNaN(value)){alert('只允许输入数字');value='';}"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))"></td>
          </tr>
		  <tr bgcolor="#e3e3e3" class="tdbg"> 
            <td width="126" height="22" bgcolor="#c0c0c0"> <div align="right"><strong>是否允许会员投稿：</strong></div></td>
            <td width="218" bgcolor="#e3e3e3"><label><input name="userdraft" type="radio" value="1" checked>是
           <input name="userdraft" type="radio" value="0">否</label></td>
          </tr>
          <tr class="tdbg"> 
            <td height="22" align="center" bgcolor="#c0c0c0">&nbsp; </td>
            <td height="22" align="center" bgcolor="#e3e3e3"> <div align="left"> 
                <input name="action" type="hidden" id="action3" value="add">				
                <input name="add" type="submit" value=" 添 加 ">
              </div></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
<%
end if
Call CloseDB(conn)%>


                                                              
                                                              
                                                              