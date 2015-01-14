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
<%Dim Action,ID,fso,Ts,i,Moldboard,Text,m_name,temp,M_logo,Data,M_,M(5)
Moldboard=Request.Form("Moldboard")
ID=strint(request.QueryString("ID"))
Call OpenConn
IF Request.Form("Submit")<>"" Then
For i=0 to 5
m(i)=strint(Request.Form("M"&i))
Temp=Temp&m(i)
IF i<>5 Then Temp=Temp&"|"
Next


Set Rs=server.createobject("adodb.recordset")
Sql="Select ID,M_Footer,M_Footerstr From DNJY_Template Where ID="&ID
Rs.Open Sql,conn,1,3
Rs(1)=Moldboard
Rs(2)=Temp
Rs.Update

set fso = CreateObject("Scripting.FileSystemObject")
set Ts = Fso.OpenTextFile(Server.Mappath("../Footer.asp"),2, True)
Ts.Write M_Footer(Moldboard,Temp)
Ts.Close:Set Ts=Nothing
Set Fso=Nothing
If Htmlhome=1 Then Call HomeHtml()'重新生成首页
response.write "<script language='javascript'>	alert('模板设置成功，并已生成新的页面\n\n请重新生成静态内容页！');location.href='javascript:history.go(-1)';</script>"
End IF
Set Rs=Conn.Execute("Select M_Footer,M_Footerstr,M_Name,M_logo From DNJY_Template Where ID="&ID)
IF Not Rs.Eof Then Data=Rs.GetRows
Rs.close:Set Rs=Nothing
Call CloseDB(conn)
M_logo=Data(3,0)
M_=Split(Data(1,0),"|")
%>
<HTML>
<HEAD>
<LINK rel="stylesheet" type="text/css" href="../css/style.css">
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<META name="GENERATOR" content="Microsoft FrontPage 6.0">
<META name="ProgId" content="FrontPage.Editor.Document">
<TITLE>尾部模板</title>
</HEAD>
<BODY background="images/background.gif">                                                                 
  <TABLE width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
    <TR>
      <TD height="20" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">尾 部 模 板 管 理
	  </FONT></TD>
    </TR>
    <TR>
      <TD bgcolor="#FFFFFF"> <BR>
   <TABLE width="98%" border="0"  align=center cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
  
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">模板代码：   </TD>
      <TD valign="top"> 当前模板：
        尾部模板footer.asp，当前模板方案：<FONT color="#ff0000"><%=Data(2,0)%></font>&nbsp;&nbsp;模板标识：<%=M_logo%><br>
		<FORM name="form1" method="post">
        &nbsp;<TEXTAREA name="moldboard" cols="85" rows="30"><%=server.HTMLEncode(Data(0,0))%></TEXTAREA>
        <BR>
        </P>        </TD>
		 <TD height="26" align="left">
        模板参数：<BR>
        <FONT color="#ff0000">{$关于本站}<BR>
		{$会员帮助}<BR>
		{$服务条款}<BR>
		{$免责声明}<BR>
		{$广告服务}<BR>
		{$联系我们}<BR>
		{$客户留言}<BR>
		{$公司名称}<BR>
		{$网站地址}<BR>
		{$程序版本}<BR>
		{$联系电话}<BR>
		{$联系手机}<BR>
		{$传真}<BR>
		{$办公地址}<BR>
		{$邮政编码}<BR>
		{$网站ICP}<BR>
		{$统计代码}</FONT></TD>
      </TR>

    <TR bgcolor="#FFFFFF"> 
      <TD height="30" align=left></TD>
      <TD height="30" colspan="2" align=left>
	  <INPUT type="submit" name="Submit" value="  确 认  ">
	  <label>
	  <input name="Submit1" type="button" id="Submit1" value="返回" onClick="location='admin_Template_Management.asp?id=<%=ID%>'"/>
	  </label>        </TD>
    </TR>
  </FORM>
</TABLE>
      <BR>      </TD>
    </TR>
  </TABLE>
</BODY>                    
                    
</HTML>                    
