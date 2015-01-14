<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Call OpenConn
dim action,id,idd,i,dname,dbpath,yname,adid,follows,fdel,iPage,pic,sql1,sql2,sql3
id=strint(request.QueryString("id"))
idd=strint(request.QueryString("idd"))
action=HtmlEncode(request.querystring("action"))
dname=HtmlEncode(request("dname"))
dbpath="../"

select case action
case "add"
set rs=server.CreateObject("adodb.recordset")
 rs.open "select * from DNJY_type where name='"&dname&"' and threeid=0 and twoid=0",conn,1,1
         if rs.EOF and rs.BOF then
           rs.close
    rs.open "select * from DNJY_type where threeid=0 and twoid=0 order by id "&DN_OrderType&"",conn,1,1
	if rs.eof and rs.bof then
	id=1
	else
    id=rs("id")+1
	end if
	rs.close
	rs.open "DNJY_type",conn,1,3
	rs.AddNew
if id=0 then
   rs("id")=1
   else
   rs("id")=id
    end If
    
rs("name")=dname
rs("indexid")=strint(request("indexid"))
rs("indexshow")=request("indexshow")
rs("titlecolor")=request("colorshow")
rs("twoid")=0
rs("threeid")=0
rs.Update
Call CloseRs(rs)
else
response.write "<script>window.alert('你输入的一级目录名称:["&dname&"]已经有了！');history.go(-1);</script>"
end if
case "edit"
   set rs=server.CreateObject("adodb.recordset")
	rs.open "select name from DNJY_type where twoid=0 and threeid=0  and name='"&dname&"' and i<>"&idd&"",conn,1,1
	if not rs.eof Then
	response.Write "<script language=javascript>alert('此分类名：["&dname&"]已经存在，请重新选择一个!');location.replace('?');</script>"	
	set rs=Nothing
    response.end	
	end If
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_type where twoid=0 and id="&id,conn,1,3
yname=rs("name")
rs("name")=dname
rs("indexid")=strint(request("indexid"))
rs("indexshow")=request("indexshow")
rs("titlecolor")=request("colorshow")
rs.Update
Call CloseRs(rs)
sql="update DNJY_info set type_one='"&dname&"' where type_oneid="&id&" "
conn.execute sql

case "del"
set rs=server.CreateObject("adodb.recordset")
sql="select * from DNJY_type where twoid>0 and id ="&id
rs.open sql,conn,3,2
if not rs.eof Then
response.Write "<script language=javascript>alert('此分类下级分类非空，请先删除下级分类才能删除本分类!');location.replace('?');</script>"
set rs=Nothing
response.End
End If
conn.execute ("delete from DNJY_type where id="&id&"")
'删除关联分类的信息及相关文件开始
sql="select * from DNJY_info where type_oneid = "&id
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,3,2
Set fdel = CreateObject("Scripting.FileSystemObject")	
if not rs.eof then
 For iPage = 1 To rs.recordcount
 adid=rs("id")
 pic=rs("tupian")
rs.delete
rs.update
sql1="delete from DNJY_Message where adid ="&adid&""    
conn.execute(sql1)
rs.movenext  
   '''''''''''删除html开始''''''''''''''''
				 if fdel.fileExists(Server.MapPath("../html/"&trim(adid)&".html")) then
                 fdel.DeleteFile(Server.MapPath("../html/"&trim(adid)&".html"))
                 end if
				 if pic<>"0" And fdel.fileExists(Server.MapPath("../UploadFiles/infopic/"&pic&"")) then
				 fdel.DeleteFile(Server.MapPath("../UploadFiles/infopic/"&pic&""))
				 end if
'''''''''''删除html结束'''''''''''''''
'删除相关回复开始
sql1="delete from [DNJY_reply] where xxid="&adid&" "
rs.open sql1,conn,1,3
'删除相关回复END
'删除相关举报开始
sql2="delete from [DNJY_Bad_info] where infoid="&adid&" "
rs.open sql2,conn,1,3
'删除相关举报END
'删除相关收藏开始
sql3="delete from [DNJY_favorites] where scid='"&adid&"' "
rs.open sql3,conn,1,3
'删除相关收藏END
             
'''''''''''删除html结束'''''''''''''''
if rs.eof then exit for
next
set fdel=nothing
Call CloseRs(rs)                                 
end if
Call CloseDB(conn)
'删除关联分类的信息及相关文件END
response.Redirect "type_Level1.asp" 
end select
%>
<HTML>
<HEAD>
<LINK rel="stylesheet" type="text/css" href="../css/style.css">
<SCRIPT language="javascript" src="admin.js"></SCRIPT>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<META name="GENERATOR" content="Microsoft FrontPage 4.0">
<META name="ProgId" content="FrontPage.Editor.Document">
<TITLE>分类管理</TITLE>
</HEAD>

<BODY background="images/background.gif">
<IFRAME width="260" height="165" id="colourPalette" src="nc_selcolor.htm" style="visibility:hidden; position: absolute; left: 0px; top: 0px;border:1px gray solid" frameborder="0" scrolling="no" ></IFRAME>

<TABLE width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
  <TR> 
    <TD height="20" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">一&nbsp;级&nbsp;分 类 栏&nbsp;目&nbsp;管 理</FONT></TD>
  </TR>
  <TR> 
    <TD bgcolor="#FFFFFF"> &nbsp;首页 <BR>
	<TABLE width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
        <TR align="center" bgcolor="#FFFFFF" height="20"> 
          <TD width="30">编号</TD>
          <TD>分类名称</TD>
          <TD>显示顺序</TD>
          <TD>首页显示</TD>
          <TD>首页配色</TD>
          <TD>管理操作</TD>
        </TR>
        <%set rs=server.CreateObject("adodb.recordset")
		  rs.Open "select * from DNJY_type where twoid=0 and id>0 order by id",conn,1,1		  
		  if rs.EOF and rs.BOF then
          response.write"<tr bgcolor=#FFFFFF><td colspan='6'><p align='center'><font color='red'>暂无分类！请先添加分类。如果觉得官方的默认分类合用，可</font><a href=""Classify_in_out.asp?Action=xxfl"" TARGET=right><FONT color=""#0000FF"" >导入官方默认信息分类。</font></a>如果不合用请勿导入，否则一个个删除非常麻烦！<a href=""http://test.ip126.com/type_list.asp"" target=_blank><FONT color=""#0000FF"" >官方默认分类参考</font></a></td></tr></table><br>"
		  follows=0
		  else
		  do while not rs.EOF
		  i=i+1
		  %>
        <FORM name="form1" method="post" action="type_Level1.asp?action=edit&id=<%=int(rs("id"))%>&idd=<%=rs("i")%>">
          <TR bgcolor="#FFFFFF" align="center">
		  <TD><%=i%></TD>
            <TD><INPUT name="dname" type="text" size="30" value="<%=rs("name")%>">
            &nbsp;[<A href="type_Level2.asp?id=<%=rs("id")%>">子目录</A>]</TD>
            <TD><INPUT name="indexid" type="text" size="3" value="<%=rs("indexid")%>"></TD>
            <TD><SELECT name="indexshow">
<OPTION value="yes" <%if rs("indexshow")="yes" then%>selected<%end if%>>显示</OPTION>
<OPTION value="no" <%if rs("indexshow")="no" then%>selected<%end if%>>隐藏</OPTION>
                    </SELECT></TD>
            <TD><INPUT name="colorshow" type="text" style="background:<%=rs("titlecolor")%>" onClick="Getcolor(this)" value="<%=rs("titlecolor")%>" size="7" maxlength="7"></TD>
            <TD><INPUT type="submit" name="Submit" value="修 改">
                &nbsp; <INPUT type="button" name="DEL" onClick="{if(confirm('确定要删除这个分类吗？\n将删除此分类下所有信息\n此操作不可以恢复！')){location.href='?name=<%=rs("name")%>&action=del&id=<%=rs("id")%>';}return false;}" value="删除" >
            </TD>
          </TR>
        </FORM>
        <%
		rs.MoveNext
          loop
          follows=rs.RecordCount
          end If%>
      </TABLE>
	<BR></TD>
  </TR>
</TABLE>
<BR>
<TABLE width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
  <TR> 
    <TD height="20" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">添 加 一&nbsp;级&nbsp;分 类 名 称</FONT></TD>
  </TR>
  <TR> 
    <TD bgcolor="#FFFFFF">
	<BR>
	<TABLE width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
        <TR align="center" bgcolor="#FFFFFF" height="20"> 
          <TD width="30">编号 </TD>
          <TD>分类名称</TD>
          <TD>显示顺序</TD>
          <TD>首页显示</TD>
          <TD>首页配色</TD>
          <TD>确定操作</TD>
        </TR>
        <FORM name="form1" method="post" action="type_Level1.asp?action=add">
          <TR align="center" bgcolor="#FFFFFF"> 
            <TD><%=rs.RecordCount+1%></TD>
	        <TD><INPUT name="dname" type="text" id="dname" size="30"></TD>
	        <TD><INPUT name="indexid" type="text" value="0" size="3"></TD>
	        <TD><SELECT name="indexshow">
                <OPTION value="yes" selected>显示</OPTION>
                <OPTION value="no">隐藏</OPTION>
            </SELECT></TD>
	         <TD><INPUT name="colorshow" type="text" style="background:#000000" onClick="Getcolor(this)" value="#000000" size="7" maxlength="7"></TD>
	        <TD><INPUT type="submit" name="Submit3" value="添 加" onClick="if(dname.value==''){alert('分类名不能为空');return false;}"></TD>
          </TR>
        </FORM>
      </TABLE>
	<BR></TD>
  </TR>
</TABLE>
</BODY>                                                                                                                          
<%Call CloseRs(rs)
Call CloseDB(conn)%>                                                                                                                           
</HTML>  
