<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("02")
Call OpenConn
dim id,mailsmtp,mailname,mailpass,mailform
    id=request("id")
	set rs=server.createobject("adodb.recordset")
	sql="select * from [DNJY_mail] where id="&cstr(id)
	rs.open sql,conn,1,3
if rs.eof or rs.bof then
response.write "没有邮件系统参数！"
Call CloseRs(rs)
response.End
End If
If request("ok")=1 And (request("mailsmtp")="" Or request("mailname")="" Or request("mailpass")="" Or request("mailform")="" Or request("maillog")="") Then
Response.Write ("<script language=javascript>alert('各个参数要认真填写!');history.go(-1);</script>")
response.end
End if
    if request("ok")=1 then
    rs("mailsmtp")=request("mailsmtp")
    rs("mailname")=request("mailname")
    rs("mailpass")=request("mailpass")
    rs("mailform")=request("mailform")
	rs("maillog")=request("maillog")
    rs.update
    response.write "<script language=JavaScript>" & chr(13) & "alert('修改成功！');" &"window.location='Mail_system.asp?id=1'" & "</script>"
    response.end
    else
%>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
}
.style1 {
	color: #FFFFFF;
	font-weight: bold;
	font-size: 12px;
}
-->
</style><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center">
      <form name="form1" method="POST" action="Mail_system.asp?ok=1&id=<%=request("id")%>">
        <table border="1" cellspacing="0" cellpadding="0" bgcolor="#F5F5F5" width="98%" style="border-collapse: collapse" bordercolor="#ADAED6">
          <tr> 
            <td height="28" bgcolor="#BDBEDE"><span class="style1">&nbsp;邮件系统配置</span><font color="#FF0000">（请将邮件系统参数换为您自己的）</font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" align="center" width="700"><TABLE>
            <TR>
				<TD width="600" colspan=4><font color="#ff0000"> 注意：由于很多邮局、尤其是免费邮箱群发邮件都受监测，可能因过密或内容相同而被暂时或永久禁止发信邮箱的发信功能，如果需经常群发邮件的建议购买不受限制的企业邮局！</font></td>
                </tr>
                <tr> 
                  <td align="right">发信人邮箱：</td>
                  <td> 
                    <input type="text" name="mailform" value="<%=rs("mailform")%>" size="20"> 作为网站管理的统一邮箱</td>
                </tr>
                <tr> 
                  <td align="right">邮箱帐号：</td>
                  <td> 
                    <input type="text" name="mailname" value="<%=rs("mailname")%>" size="20">
                   邮箱帐号一般为@前面部分</td>
                </tr>
                <tr> 
                  <td align="right">访问密码：</td>
                  <td> 
                    <input type="password" name="mailpass" value="<%=rs("mailpass")%>" size="20"> 邮箱的密码</td>
                </tr>
                <tr> 
                  <td width="20%" align="right">Smtp服务器：</td>
                  <td width="80%"> 
                    <input type="text" name="mailsmtp" value="<%=rs("mailsmtp")%>" size="20"> 
                  根据邮箱的smtp服务器格式填写<br>（如网易的163邮箱为smtp.163.com，搜狐邮箱为smtp.sohu.com，新浪邮箱为smtp.sina.com）</td>
                </tr>
                <tr> 
                  <td width="20%" align="right">邮件日志保留：</td>
                  <td width="80%"> 
                    <input type="text" name="maillog" value="<%=rs("maillog")%>" size="20" onKeyUp="if(isNaN(value)){alert('只允许输入数字');value='';}" onafterpaste="this.value=this.value.replace(/\D/g,'')">  天<br>
                  （邮件系统发送邮件的日志保留天数，0为不删除，为节省资源，建议设置一定时间自动删除）</td>
                </tr>
              </table></TD>
				<TD width="200">
<%On Error Resume Next 
    Err = 0
Dim JMails,getvers 
Set JMails=Server.CreateObject("JMail.Message")
If JMails Is Nothing Then    
    Response.Write "服务器邮件发送JMail组件检测结果：<br><font color=red><b>×</b></font> <font color=#555555>不支持，请安装JMail组件，否则无法发送邮件！</font>" 
Else    
If 0 = Err Then getvers=JMails.version
    Response.Write "服务器邮件发送JMail组件检测结果：<br><font color=#0000ff><b>√</b></font>  版本 "&getvers 
End If 
Set JMails = Nothing
    Err = 0%></TD>


            </TR>
            </TABLE> 
             
            </td>
          </tr>
          <tr>
            <td height="35" align="center" bgcolor="#eeeeee"> 
              <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td><div align="center">
                    <input type="submit" name="Submit" value="确认设置" class="Tips_bo">
                  </div></td>
                  <td><div align="center">
                    <input type="reset" name="Submit" value="取消设置" class="Tips_bo">
                  </div></td>
                </tr>
              </table>
            </td>
          </tr>
         
        </table>
<%end if
Call CloseRs(rs)
Call CloseDB(conn)%>
      </form>
    </td>
  </tr>
</table>
