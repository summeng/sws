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
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<script language=javascript>
<!--
function CheckAll(form){
for (var i=0;i<form.elements.length;i++){
var e = form.elements[i];
if (e.name != 'chkall') e.checked = form.chkall.checked; 
}
}
-->
</script>
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
<style type="text/css">
/*��ʾ��Ϣ*/
.info {position:relative;background:#;color:#666; text-decoration:none;font-size:12px;width:0px;height:0px;text-align:center;border:0px solid #ccc;line-height:18px;}/*�������ӵ�����,һ��Ҫ����Ϊrelative����ʹ��ʾ�����������*/
.info:hover {background:#eee;color:#333;}
.info span {display: none }/*���������µ�spanΪ����״̬*/
.info:hover span /*����hover�µ�span����Ϊ����״̬,��������ʾ���λ��*/{display:block;position:absolute;top:0px;left:25px;width:200px;
border:1px solid #DC5826; background:#FFF3B4; color:#333;padding:5px;text-align:left;}
</style>
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
	if (document.email.email.value=="" )
	{
	  alert("��������")
	  document.email.email.focus()
	  return false
	 }
    if((!checkemail(email.email.value))&&(document.email.email.value!=""))
	{
    alert("������Email��ַ����ȷ������������!");
    document.email.email.focus();
    return false;
    }
if (document.email.mail_subject.value=="" )
	{
	  alert("�����ʼ�����")
	  document.email.mail_subject.focus()
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

<%dim action,email,aJMail,selectedid,i,thisEmail,sssstemp,C_M_Chk
Set aJMail=Server.CreateObject("JMail.Message")   
If aJMail Is Nothing Then  
response.write "<script LANGUAGE='javascript'>alert('��ܰС��ʾ����������֧��JMail.Message�������������װ�����������������ʼ�����!');history.go(-1);</script>"
Response.End
End   If
If trim(request("userqf"))="ok" And trim(request("selectedid"))="" Then
response.write "<script LANGUAGE='javascript'>alert('��ܰС��ʾ��δѡ���Ա��֪�ʼ�Ҫ����˭!');history.go(-1);</script>"
Response.End
End   If
action=request("action")
if action="send" Then
call sendmail()
else
email=request.ServerVariables("query_string")'������־����
selectedid=trim(request("selectedid"))
If selectedid<>"" Then'���Ի�Ա����
selectedid=split(selectedid,",")
Call OpenConn
for i=0 to ubound(selectedid)
set rs=server.createobject("adodb.recordset")
sql="select email from [DNJY_user] where username='"&trim(cstr(selectedid(i)))&"'"
rs.open sql,conn,1,1
thisEmail=rs(0)
if instr(sssstemp,thisEmail)=0 Then
sssstemp=sssstemp&thisEmail
If i<>ubound(selectedid) Then sssstemp=sssstemp&";"
end If
Next
email=sssstemp
Call CloseRs(rs)
Call CloseDB(conn)
End if%>
<form action="sendmail.asp" method=post name="email" onSubmit="return formcheck()">
<table width="100%" border="1"  style="border-collapse: collapse; border-style: dotted; border-width: 0px" bordercolor="#333333" cellspacing="0" cellpadding="2" style="word-break:break-all">
<tr><td colspan=4 bgcolor="#BDBEDE" height=28 align=center><B>���Ա���͵�����־</b></td></tr>
<tr><td width=20% align=center colspan=2><font color="#ff0000"> ע�⣺���ںܶ��ʾ֡��������������Ⱥ���ʼ����ܼ�⣬��������ܻ�������ͬ������ʱ�����ý�ֹ���Ź��ܣ�����辭��Ⱥ���ʼ��Ľ��鹺�������Ƶ���ҵ�ʾ֣�</font></td></TR>
<tr><td width=20%  align="right">�� �� �䣺</td>
<td width="650"><TEXTAREA NAME="email" rows=6 cols="90" style="width=550 ; overflow:auto;" cols="50"><%=email%></TEXTAREA><br><a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>��д�ռ��������ַ��������԰�Ƕ��Ż�ֺŸ���ǧ��Ҫ�������ʼ�Ŷ~���������</span></a> </td></TR>
<tr><td width="20%"  align="right">�ʼ����⣺</td><td width="600"><input type=text name="mail_subject" style="width=90%;" maxlength=50 size="70"><br> <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>��д�ʼ����⼼�ɣ��õ����������˵���Ķ����������ܻᱻֱ��ɾ��</span></a></td></TR>
	<tr>                                                                   
    <td width=20%  align="right">�ʼ���ʽ��</td>                                                                   
    <td width="86%" height="26"><input type="radio" name="radiobutton" value="text" checked style="border:1px black solid">                                                                   
      �ı��ʼ�&nbsp; <input type="radio" name="radiobutton" value="html" style="border:1px black solid">                                                                          
        HTML�ʼ�&nbsp;&nbsp;<font color="#ff0000"> �����������ݿ�Ҫ����ѡ����ʼ���ʽ��Ӧ��д������ͬʱ��д��</font></td>                                                                           
  </tr> 
<tr><td width=20%  align="right">�ı���ʽ�ʼ����ݣ�<br><br><font color=red>��1000�ַ����ڣ�</font></td><td><TEXTAREA NAME="mailbody" ROWS="15" cols="75"style="width=550px ; overflow:auto;"></TEXTAREA><br> <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>֧���ı���ʽ<br>���ռ��˵����䲻֧��html��ѡ��˷�ʽ</span></a> </td></TR>
<tr><td width=20%  align="right">HTML��ʽ�ʼ����ݣ�<br><br><font color=red>��1000�ַ����ڣ�</font></td><td>
<textarea id="editor" name="mailbodyhtm"  style="width:550px;height:280px;visibility:visible"></textarea>
&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>֧��html�����ռ��˵����䲻֧��html�򿴵�����һ�Ѵ������</span></a></td></TR>
 
  <tr>                                                                          
    <td width=20%  align="right">�ʼ��ȼ���</td>                                                                          
    <td width="86%" height="26"><input type="radio" name="dengji" value="1" style="border:1px black solid">                                                                          
      �Ӽ�&nbsp; <input type="radio" name="dengji" value="3" checked style="border:1px black solid">                                                                          
      ��ͨ&nbsp; <input type="radio" name="dengji" value="5" style="border:1px black solid">                                                                          
      �ͼ�&nbsp;&nbsp;</td>                                                                          
  </tr>
<tr><td colspan=2 bgcolor='#DADAE9' align=center>
<input type=hidden name=action value="send">
<input type=submit value="��������" >
&nbsp; &nbsp; &nbsp; <input type="reset" value="������д">
</table>
</form>
<%dim mail_subject,mailbody,mailserver,mailpassword,sitename
end if
sub sendmail()
Dim x,io,sendok,i
If Request("mailbody") <> "" and Request("mailbodyhtm") <> "" Then  
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
mail_subject=trim(request("mail_subject"))
If Request("mailbody") <> "" then
mailbody=Request("mailbody")
Else
mailbody=Request("mailbodyhtm")
End if
email=trim(request("email"))
email=replace(email," ","")
email=replace(email,"��",";")
email=replace(email,"��",";")
email=replace(email,",",";")
if email="" or mail_subject="" or mailbody="" then
response.write "<script language='javascript'>"
response.write "alert('�����ˣ��ռ��ˡ��ʼ����⡢�ʼ����ݱ���������д��');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if

if UBound(split(email,";"))>20 then
response.write "<script language='javascript'>"
response.write "alert('�����ˣ�����д���ռ���̫����,��20����~~');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if

mailserver=mailsmtp		
mailname=mailname			
mailpassword=mailpass
sitename=webname
mailbody=mailbody+"����ӭ����"+webname+"http://"+web
x=split(email,";")

Dim aJMail   
Set aJMail=Server.CreateObject("JMail.Message")   
If aJMail Is Nothing Then  
response.write "<script LANGUAGE='javascript'>alert('��ܰС��ʾ����������֧��JMail.Message�������������װ�����������������ʼ�����!');history.go(-1);</script>"
Response.End
End   If

Dim JMail,ErrStr
for io=0 to Ubound(x)
If ChkMail(x(io)) =true Then'�ж������Ƿ�Ϸ�
'jmail�ʼ���������ʼ���ʼ�������ֵҪ��ȡ��
Set jmail = Server.CreateObject("JMail.Message") 
jmail.silent = true '����������󣬷���FALSE��TRUE��ֵ
jmail.Charset = "gb2312"  '�ʼ������ֱ���Ϊ����
If Request.form("radiobutton") = "html" Then
jmail.ContentType = "text/html"    '�ʼ��ĸ�ʽΪHTML��ʽ
end if
jmail.MailServerUserName = mailname '����smtp��������֤��½�� ���ʾ����κ�һ���û���Email��ַ�� 
jmail.MailServerPassword = mailpassword '����smtp��������֤���� ���û�Email�ʺŶ�Ӧ�����룩 
jmail.From = mailform '���������䣬��������������ʼ�������������뷢�ŷ�����ͬ���û������䣡 
jmail.FromName = webname '���������� 
jmail.AddRecipient(x(io)) '�ռ���Email 
jmail.Subject = mail_subject '�ʼ��ı��� 
jmail.Body = mailbody '�ʼ�������
JMail.Priority =1' Request.form("dengji")      '�ʼ��Ľ������� 
jmail.Send (mailserver) 'ִ���ʼ�����smtp��������ַ����ҵ�ʾֵ�ַ��
ErrStr =jmail.errormessage
jmail.close 
set jmail = nothing '�رն���
end if'�ж������Ƿ�Ϸ�
next  

if ErrStr="" Then
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
response.write "<script language='javascript'>alert('��ϲ���ʼ��Ѿ�������');location.href='sendmail_user.asp';</script>"
Response.End
else
Response.Write "<script>alert('�����ʼ�ʧ�ܣ�\n\nʧ��ԭ��"&ErrStr&"��\n\n���ȷ�����أ�');history.back();</Script>"
Response.End 
end if
end sub
%>
<!--#include virtual="/Include/bottom_superadmin.asp" -->