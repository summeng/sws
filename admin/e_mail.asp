<!--#include file="../conn.asp"-->
<!--#include file=../config.asp-->
<!--#include file=../Include/mail.asp-->
<!--#include file=cookies.asp-->

<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("08")%>
<html>
<head>
<title>���͵�����־</title>
<meta http-equiv="Ŀ¼����" content="�ı�/html; �ַ���=gb2312">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<!--kindeditor�༭��-->
		<link rel="stylesheet" href="../KindEditor/themes/default/default.css" />
		<script charset="utf-8" src="../KindEditor/kindeditor-min.js"></script>
		<script charset="utf-8" src="../KindEditor/lang/zh_CN.js"></script>
		<script>
			var editor;			
			KindEditor.ready(function(K) {
				editor = K.create('textarea[name="mailbodyhtm"]', {
					resizeType : 2,//0:����kindeditor�༭����С���ɸı�; 1:ֻ�ܸñ�߶�; 2:���Ըı�߶ȺͿ��
					afterBlur: function(){this.sync();},//ʧȥ����ִ��this���ֵ
					allowPreviewEmoticons : false,
					allowImageUpload : false,					
					items : [
						'fontname', 'fontsize', '|', 'forecolor', 'hilitecolor', 'bold', 'italic', 'underline',
						'removeformat', '|', 'justifyleft', 'justifycenter', 'justifyright', 'insertorderedlist',
						'insertunorderedlist', '|', 'emoticons', 'image', 'link']
				});
			});
		</script>
<!--kindeditor�༭��END-->
<script language="javascript">
<!--
//��֤������ȷ��
function checkemail(email){
var str=email;
//��JavaScript�У�������ʽֻ��ʹ��"/"��ͷ�ͽ���������ʹ��˫����
var Expression=/\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*/; 
var objExp=new RegExp(Expression);
if(objExp.test(str)==true){
return true;
}else{
return false;
}
} 

//˵���ֽ������ƣ����������=========
String.prototype.Length=function(){
temp=this.replace(/([^\x00-\xff])/g,"$1$1");
return temp.length
}

function formcheck(){
    if((!checkemail(sendmail.tomail.value))&&(document.sendmail.tomail.value!=""))
	{
    alert("������Email��ַ����ȷ������������!");
    document.sendmail.tomail.focus();
    return false;
    }
if (document.sendmail.mailsubject.value=="" )
	{
	  alert("�����ʼ�����")
	  document.sendmail.mailsubject.focus()
	  return false
	 }
if (document.sendmail.mailbody.value.length>5000)  //�ֽ������ƣ���������ϡ�
     {
	 alert("�ʼ����ݳ���Ҫ��<%=5000%>�ֽ�(<%=2500%>����)�ڣ����������룡");
	  return false
  }

	  
}
// --></script>
</head>
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
.style2 {font-size: 12px}
-->
</style>
<center>
<%Dim aJMail   
Set aJMail=Server.CreateObject("JMail.Message")   
If aJMail Is Nothing Then  
response.write "<script LANGUAGE='javascript'>alert('��ܰС��ʾ����������֧��JMail.Message�������������װ�����������������ʼ�����!');history.go(-1);</script>"
Response.End
End If
If Request.form("tomail") <> "" and Request("qunfa") <> "" Then  
response.write "<script LANGUAGE='javascript'>alert('��ܰС��ʾ����д�ռ����������ͬʱѡ��Ⱥ��!');history.go(-1);</script>"
Response.End
End If
If Request.form("mailbody") <> "" and Request("mailbodyhtm") <> "" Then  
response.write "<script LANGUAGE='javascript'>alert('��ܰС��ʾ�����ָ�ʽ�ʼ����ݲ���ͬʱ��д!');history.go(-1);</script>"
Response.End
End If
If Request("radiobutton") = "text" and Request("mailbody")="" Then  
response.write "<script LANGUAGE='javascript'>alert('��ܰС��ʾ��ѡ���ı���ʽ�ʼ��뽫������д���ı���ʽ�ʼ����ݿ�!');history.go(-1);</script>"
Response.End
End If
If Request("radiobutton") = "html" and Request("mailbodyhtm")="" Then  
response.write "<script LANGUAGE='javascript'>alert('��ܰС��ʾ��ѡ��HTML��ʽ�ʼ��뽫������д��HTML��ʽ�ʼ����ݿ�!');history.go(-1);</script>"
Response.End
End If
If Request.form("mailbody") <> "" then
mailbody=Request.form("mailbody")
Else
mailbody=Request.form("mailbodyhtm")
End if
dim userid,rs1,tomail,mailsubject,mailbody,host,dengji,ErrStr,jmail,sitename,email,i

if Request.form("action")="send" then
call mailto()
Else
call mailedit()
End if
sub mailto()
If ChkMail(Request.form("tomail")) =False Then
response.Write "<script language=javascript>alert('�����������ʽ����!');history.go(-1);</script>"
response.end
End If
if mailbody <> "" and Request.form("mailsubject") <> "" then 
userid=request.querystring("id")
set rs1=server.createobject("adodb.recordset")
rs1.open "select email from DNJY_user where maillist=1 and email<>'' " ,conn,1,1	
	email = Request.form("tomail") 
	mailsubject = Request.form("mailsubject") '�ż�����
	mailbody = mailbody '�ż�����
	host = ""
	dengji = Request.form("dengji")  '�ʼ��ȼ�
	ErrStr = "ture"	

if Request.form("qunfa") = "" then'�����Ƿ�Ⱥ��
'jmail�ʼ���������ʼ���ʼ�������ֵҪ��ȡ��
Set jmail = Server.CreateObject("JMail.Message") 
jmail.silent = true '����������󣬷���FALSE��TRUE��ֵ
jmail.Logging = true '�����ʼ���־
jmail.Charset = "gb2312"  '�ʼ������ֱ���Ϊ����
If Request("radiobutton") = "html" Then
jmail.ContentType = "text/html"    '�ʼ��ĸ�ʽΪHTML��ʽ
end if
jmail.MailServerUserName = mailname '����smtp��������֤��½�� ���ʾ����κ�һ���û���Email��ַ�� 
jmail.MailServerPassword = mailpass '����smtp��������֤���� ���û�Email�ʺŶ�Ӧ�����룩 
jmail.From = mailform '���������䣬��������������ʼ�������������뷢�ŷ�����ͬ���û������䣡 
jmail.FromName = webname '���������� 
jmail.AddRecipient email '�ռ���Email 
jmail.Subject = mailsubject '�ʼ��ı��� 
jmail.Body = mailbody '�ʼ�������
JMail.Priority = dengji      '�ʼ��Ľ������� 
jmail.Send (mailsmtp) 'ִ���ʼ�����smtp��������ַ����ҵ�ʾֵ�ַ��
ErrStr =jmail.errormessage
jmail.close 
set jmail = nothing '�رն���
'�ʼ����ͽ���


else'Ⱥ��
for i=1 to rs1.recordcount
email = rs1("email")  '�ͻ�����
If ChkMail(email) =true Then'�ж������Ƿ�Ϸ�
'jmail�ʼ���������ʼ���ʼ�������ֵҪ��ȡ��
Set jmail = Server.CreateObject("JMail.Message") 
jmail.silent = true '����������󣬷���FALSE��TRUE��ֵ
jmail.Charset = "gb2312"  '�ʼ������ֱ���Ϊ����
If Request("radiobutton") = "html" Then
jmail.ContentType = "text/html"    '�ʼ��ĸ�ʽΪHTML��ʽ
end if
jmail.MailServerUserName = mailname '����smtp��������֤��½�� ���ʾ����κ�һ���û���Email��ַ�� 
jmail.MailServerPassword = mailpass '����smtp��������֤���� ���û�Email�ʺŶ�Ӧ�����룩 
jmail.From = mailform '���������䣬��������������ʼ�������������뷢�ŷ�����ͬ���û������䣡 
jmail.FromName = webname '���������� 
jmail.AddRecipient email '�ռ���Email 
jmail.Subject = mailsubject '�ʼ��ı��� 
jmail.Body = mailbody '�ʼ�������
JMail.Priority = dengji      '�ʼ��Ľ������� 
jmail.Send (mailsmtp) 'ִ���ʼ�����smtp��������ַ����ҵ�ʾֵ�ַ��
ErrStr =jmail.errormessage
jmail.close 
set jmail = nothing '�رն���
'�ʼ����ͽ���
end if'�ж������Ƿ�Ϸ�
		rs1.movenext
		next
		rs1.close
	end if	'Ⱥ���жϽ���
if ErrStr = "" then
	Response.Write "��ϲ���ʼ����ͳɹ���" 
'����������ͬʱд�ʼ���־��ʼ��������
Dim rslog
set rslog=server.createobject("adodb.recordset")
sql="select * from DNJY_log"
rslog.open sql,conn,1,3
rslog.addnew
rslog("username")=request.cookies("administrator")
rslog("outbox")=mailform
rslog("inbox")=email
rslog("title")=mailsubject
rslog("content")=mailbody
rslog("Sendtime")=now()
rslog("lock")=0
rslog.update
rslog.close
set rslog=nothing
'��������ͬʱд�ʼ���־END��������		
else
	
	Response.Write "���ź����ʼ�δ�ܷ��ͣ�<br>�����ʼ�ϵͳ�������á��Ƿ��ж�����־��Ա."
end if	'���ͽ���жϽ���     
Response.end
else		     
response.write "<script language='javascript'>"
response.write "alert('�����ˣ��ʼ����⡢�ʼ����ݱ���������д��');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end                                                                          
end if  '�Ƿ����жϽ���                                                                         
Call CloseDB(conn)
end Sub
sub mailedit()%> 
<table border="1" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%" cellpadding="0">
  <tr bgcolor="#BDBEDE">
    <td width="100%" height="28" colspan="8" align="center"><b><span class="style2">���͵�����־</span></b></td>
  </tr>
</table>
  <form name="sendmail" action="e_mail.asp" method="post" onSubmit="return formcheck()">
 <table border="1" width="95%" height="161"  background="images/greystrip.gif"  bordercolordark="#FFFFFF"  cellspacing="0">
      <tr>
    <td width="100%" colspan="2" height="20">
      <p align="center" ><font color="#ff0000"> ע�⣺���ںܶ��ʾ֡��������������Ⱥ���ʼ����ܼ�⣬��������ܻ�������ͬ������ʱ�����ý�ֹ���Ź��ܣ�����辭��Ⱥ���ʼ��Ľ��鹺�������Ƶ���ҵ�ʾ֣�</font><br>��־Ⱥ�����ն���Ϊ�����˵�����־�Ļ�Ա��ָ��������⣩</p>
    </td>
  </tr>
  <tr>
    <td width="100%" colspan="2" height="20" background="images/mmto.gif">
      <p align="center"></p>
    </td>
  </tr>
  <tr>
    <td width="14%" height="26" align="right">�������䣺</td>
    <td width="86%" height="26"> <input type="text" name="tomail" style="border:1px black solid" size="36" value="<%=Request.QueryString("mail")%>">&nbsp;                                                                    
      �Ƿ�Ⱥ����<input type="checkbox" name="qunfa" value="ON" <%if Request.QueryString("mail") = "" then Response.Write "checked"%>>                                                                    
    �������ָ�����䣬��Ĭ��������־�Ļ�Ա����Ⱥ���ʼ���</td>                                                                   
  </tr>                                                                   
  <tr>                                                                   
    <td width="14%" height="26" align="right">�ʼ����⣺</td>                                                                   
    <td width="86%" height="26"><input type="text" name="mailsubject" style="border:1px black solid" size="36">                                                                    
    </td>                                                                   
  </tr>                                                                   
  <tr>                                                                   
    <td width="14%" height="26" align="right">�ʼ���ʽ��</td>                                                                   
    <td width="86%" height="26"><input type="radio" name="radiobutton" value="text" checked  style="border:1px black solid">                                                                   
      �ı��ʼ�&nbsp; <input type="radio" name="radiobutton" value="html" style="border:1px black solid">                                                                          
        HTML�ʼ�&nbsp;<font color="#ff0000"> �����������ݿ�Ҫ����ѡ����ʼ���ʽ��Ӧ��д������ͬʱ��д��</font></td>                                                                           
  </tr>                                                                           
  <tr>                                                                           
    <td width="14%" height="26" align="right">�ʼ����ݣ�<br>���ı���ʽ��</td>                                                                          
    <td width="86%" height="26"><textarea name="mailbody" cols="90" rows="15" style="border:1px black solid"></textarea>  </td>                                                                          
  </tr> 
  <tr>                                                                           
    <td width="14%" height="26" align="right">�ʼ����ݣ�<br>��HTML��ʽ��</td>                                                                          
    <td width="86%" height="26"><textarea id="editor" name="mailbodyhtm"  style="width:550px;height:280px;visibility:visible"></textarea></td>                                                                          
  </tr>   
  <tr>                                                                          
    <td width="14%" height="26" align="right">�ʼ��ȼ���</td>                                                                          
    <td width="86%" height="26"><input type="radio" name="dengji" value="1" style="border:1px black solid">                                                                          
      �Ӽ�&nbsp; <input type="radio" name="dengji" value="3" checked style="border:1px black solid">                                                                          
      ��ͨ&nbsp; <input type="radio" name="dengji" value="5" style="border:1px black solid">                                                                          
      �ͼ�&nbsp;&nbsp;</td>                                                                          
  </tr>                                                                          
  <tr>                                                                          
    <td width="100%" height="26" colspan="2">                                                                          
      <p align="center"><input type="submit" name="send" value="����" style="border-style: solid; border-width: 1">                                                                           
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="reset" name="Submit" value="ȡ��" style="border-style: solid; border-width: 1">     <input type=hidden name="action" value="send">                                                                     
    </td>                                                                          
  </tr>                                                                          
 </table></form> 
 <%end Sub%>
</center>                                                                              
</body>                                                                              
</html>
<!--#include virtual="/Include/bottom_superadmin.asp" -->