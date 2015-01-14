<!--#include file="../conn.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file="../Include/Form_board .asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Dim Action,ID,fso,Ts,i,Moldboard,Text,m_name,temp,M_dir,M_cssw,M_logo,Data,M_,M(5)
Moldboard=Request.Form("Moldboard")
ID=strint(request.QueryString("ID"))
Call OpenConn
IF Request.Form("Submit")<>"" Then

Set Rs=server.createobject("adodb.recordset")
Sql="Select ID,M_css,M_dir From DNJY_Template Where ID="&ID
Rs.Open Sql,conn,1,3
Rs(1)=Moldboard
M_dir=Rs(2)
Rs.Update
Rs.close
set fso = CreateObject("Scripting.FileSystemObject")
set Ts = Fso.OpenTextFile(Server.Mappath("../skin/"&M_dir&"/Style.css"),2, True)
Ts.Write M_css(Moldboard)
Ts.Close:Set Ts=Nothing
Set Fso=Nothing
'If Htmlhome=1 Then Call HomeHtml()'重新生成首页
response.write "<script language='javascript'>	alert('模板"&M_Name&"的CSS设置成功，并已生成新的Style.css文件！');location.href='javascript:history.go(-1)';</script>"
End If

set rs=server.CreateObject("adodb.recordset")
rs.Open "select M_dir,M_Name,M_css,M_logo from DNJY_Template where id="&id&" order by id",conn,1,1
%>
<HTML>
<HEAD>
<LINK rel="stylesheet" type="text/css" href="../css/style.css">
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<META name="GENERATOR" content="Microsoft FrontPage 6.0">
<META name="ProgId" content="FrontPage.Editor.Document">
<TITLE>模板<%=Rs(1)%>的CSS风格</title>
</HEAD>
<BODY background="images/background.gif">                                                                 
  <TABLE width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
    <TR>
      <TD height="20" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">模板<%=Rs(1)%>的CSS风格管理
	  </FONT></TD>
    </TR>
    <TR>
      <TD bgcolor="#FFFFFF"> <BR>
   <TABLE width="98%" border="0"  align=center cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
  
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">CSS风格代码：   </TD>
      <TD valign="top"> 当前CSS风格<%=strInstallDir%>skin/<%=Rs(0)%>/style.css，当前模板方案：<FONT color="#ff0000"><%=Rs(1)%></font>&nbsp;&nbsp;模板标识：<%=Rs(3)%><br>
		<FORM name="form1" method="post">
        &nbsp;<TEXTAREA name="Moldboard" cols="100" rows="30"><%=Rs(2)%></TEXTAREA>
        <BR>
        </P>        </TD>
		 <TD height="26" align="left">
        </TD>
      </TR>

    <TR bgcolor="#FFFFFF"> 
      <TD height="30" align=left></TD>
      <TD height="30" colspan="2" align=left>
	  <INPUT type="submit" name="Submit" value="  确 认  ">
	  <label>
	  <input name="Submit1" type="button" id="Submit1" value="返回" onClick="locatio n='admin_Template_Management.asp?id=<%=ID%>'"/>
	  </label>        </TD>
    </TR>
  </FORM>
</TABLE>
      <BR>      </TD>
    </TR>
  </TABLE>
</BODY>                    
                    
</HTML>                    
<%
Rs.close:Set Rs=Nothing
Call CloseDB(conn)
%>