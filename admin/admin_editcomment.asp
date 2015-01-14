<!--#include file="../conn.asp"-->
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
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
<BODY>
<script language=javascript src=../Include/mouse_on_title.js></script>

 <BODY>
<%Call OpenConn
Dim id,rs2,pinglunname,pingluncontent
id=cint(request.querystring("id"))
if request("action")="save" then
	pinglunname=Replace(Request.Form("pinglunname"),"'","＇") 
	pingluncontent=Replace(Request.Form("pingluncontent"),"'","＇") 
	set rs=server.createobject("adodb.recordset")
	rs.open "select * from DNJY_news_pinglun where pinglunid="&id&"",conn,1,3
	rs("pinglunname")=replace(replace(replace(replace(pinglunname,"'","‘"),"<","&lt;"),">","&gt;")," ","&nbsp;")
	rs("pingluncontent")=replace(replace(replace(replace(pingluncontent,"'","‘"),"<","&lt;"),">","&gt;")," ","&nbsp;")
	rs.update
	rs.close
	set rs=Nothing
    response.Write "<script language=javascript>alert('已成功修改!');location.replace('admin_Comment.asp');</script>"
	response.end
end if%>
<%set rs2=server.CreateObject("adodb.recordset") 
rs2.open "select * from DNJY_news_pinglun  where pinglunid="&request.querystring("id")&"",conn,1,1
%>

  		  <table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#ffffff">
<tr class=backs><td align="center" class=td height=18 colspan="9"><font color="#FFFFFF">编辑新闻评论</font></td>
		  <form name="pinglunform" method="post" action="?id=<%=rs2("pinglunid")%>&action=save">
            <tr bgcolor="#f5f5f5">
              <td width="20%" align="left">&nbsp;评论人姓名：</td>
              <td width="80%" align="left"><input name="pinglunname" type="text" id="pinglunname" size="20" maxlength="8" value=<%=trim(rs2("pinglunname"))%>></td>
            </tr>
            <tr bgcolor="#f5f5f5">
              <td align="left">&nbsp;评论正文：</td>
              <td align="left"><textarea name="pingluncontent" cols="35" rows="4" id="pingluncontent"><%=trim(rs2("pingluncontent"))%> </textarea></td>
            </tr>

            <tr bgcolor="#f5f5f5">
              <td colspan="2"><div align="center">
                  <input type="submit" name="submit" value="提 交" onClick="return check();">
                &nbsp;
                <input type="reset" name="submit2" value="清 除">
              </div></td>
            </tr></form>
          </table>	  
 </BODY>
</HTML>
<%
Call CloseRs(rs2)
Call CloseDB(conn)%>