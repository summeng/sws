<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
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
response.write "û���ʼ�ϵͳ������"
Call CloseRs(rs)
response.End
End If
If request("ok")=1 And (request("mailsmtp")="" Or request("mailname")="" Or request("mailpass")="" Or request("mailform")="" Or request("maillog")="") Then
Response.Write ("<script language=javascript>alert('��������Ҫ������д!');history.go(-1);</script>")
response.end
End if
    if request("ok")=1 then
    rs("mailsmtp")=request("mailsmtp")
    rs("mailname")=request("mailname")
    rs("mailpass")=request("mailpass")
    rs("mailform")=request("mailform")
	rs("maillog")=request("maillog")
    rs.update
    response.write "<script language=JavaScript>" & chr(13) & "alert('�޸ĳɹ���');" &"window.location='Mail_system.asp?id=1'" & "</script>"
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
            <td height="28" bgcolor="#BDBEDE"><span class="style1">&nbsp;�ʼ�ϵͳ����</span><font color="#FF0000">���뽫�ʼ�ϵͳ������Ϊ���Լ��ģ�</font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" align="center" width="700"><TABLE>
            <TR>
				<TD width="600" colspan=4><font color="#ff0000"> ע�⣺���ںܶ��ʾ֡��������������Ⱥ���ʼ����ܼ�⣬��������ܻ�������ͬ������ʱ�����ý�ֹ��������ķ��Ź��ܣ�����辭��Ⱥ���ʼ��Ľ��鹺�������Ƶ���ҵ�ʾ֣�</font></td>
                </tr>
                <tr> 
                  <td align="right">���������䣺</td>
                  <td> 
                    <input type="text" name="mailform" value="<%=rs("mailform")%>" size="20"> ��Ϊ��վ�����ͳһ����</td>
                </tr>
                <tr> 
                  <td align="right">�����ʺţ�</td>
                  <td> 
                    <input type="text" name="mailname" value="<%=rs("mailname")%>" size="20">
                   �����ʺ�һ��Ϊ@ǰ�沿��</td>
                </tr>
                <tr> 
                  <td align="right">�������룺</td>
                  <td> 
                    <input type="password" name="mailpass" value="<%=rs("mailpass")%>" size="20"> ���������</td>
                </tr>
                <tr> 
                  <td width="20%" align="right">Smtp��������</td>
                  <td width="80%"> 
                    <input type="text" name="mailsmtp" value="<%=rs("mailsmtp")%>" size="20"> 
                  ���������smtp��������ʽ��д<br>�������׵�163����Ϊsmtp.163.com���Ѻ�����Ϊsmtp.sohu.com����������Ϊsmtp.sina.com��</td>
                </tr>
                <tr> 
                  <td width="20%" align="right">�ʼ���־������</td>
                  <td width="80%"> 
                    <input type="text" name="maillog" value="<%=rs("maillog")%>" size="20" onKeyUp="if(isNaN(value)){alert('ֻ������������');value='';}" onafterpaste="this.value=this.value.replace(/\D/g,'')">  ��<br>
                  ���ʼ�ϵͳ�����ʼ�����־����������0Ϊ��ɾ����Ϊ��ʡ��Դ����������һ��ʱ���Զ�ɾ����</td>
                </tr>
              </table></TD>
				<TD width="200">
<%On Error Resume Next 
    Err = 0
Dim JMails,getvers 
Set JMails=Server.CreateObject("JMail.Message")
If JMails Is Nothing Then    
    Response.Write "�������ʼ�����JMail����������<br><font color=red><b>��</b></font> <font color=#555555>��֧�֣��밲װJMail����������޷������ʼ���</font>" 
Else    
If 0 = Err Then getvers=JMails.version
    Response.Write "�������ʼ�����JMail����������<br><font color=#0000ff><b>��</b></font>  �汾 "&getvers 
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
                    <input type="submit" name="Submit" value="ȷ������" class="Tips_bo">
                  </div></td>
                  <td><div align="center">
                    <input type="reset" name="Submit" value="ȡ������" class="Tips_bo">
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
