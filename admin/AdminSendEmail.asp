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
<%Dim Submit,aJMail,mailbody,MailTitle,UseHtml,Target,thisEmail,sssstemp,emails,SplitEmail,CanHTML,i,tomail,url_return,C_M_Chk
Set aJMail=Server.CreateObject("JMail.Message")   
If aJMail Is Nothing Then  
response.write "<script LANGUAGE='javascript'>alert('��ܰС��ʾ����������֧��JMail.Message�������������װ�����������������ʼ�����!');history.go(-1);</script>"
Response.End
End If
If Request("mailbody") <> "" and Request("mailbodyhtm") <> "" Then  
response.write "<script LANGUAGE='javascript'>alert('��ܰС��ʾ�����ָ�ʽ�ʼ����ݲ���ͬʱ��д!');history.go(-1);</script>"
Response.End
End If
If Request("UseHtml") = "text" and Request.form("mailbody")="" Then  
response.write "<script LANGUAGE='javascript'>alert('��ܰС��ʾ��ѡ���ı���ʽ�ʼ��뽫������д���ı���ʽ�ʼ����ݿ�!');history.go(-1);</script>"
Response.End
End If
If Request("UseHtml") = "HTML" and Request("mailbodyhtm")="" Then  
response.write "<script LANGUAGE='javascript'>alert('��ܰС��ʾ��ѡ��HTML��ʽ�ʼ��뽫������д��HTML��ʽ�ʼ����ݿ�!');history.go(-1);</script>"
Response.End
End If
If Request("Target")="custom" and Request("usermails") = "" Then  
response.write "<script LANGUAGE='javascript'>alert('��ܰС��ʾ��ѡ���Զ������������д����!');history.go(-1);</script>"
Response.End
End If

If Request.form("mailbody") <> "" then
mailbody=Request.form("mailbody")
Else
mailbody=Request.form("mailbodyhtm")
End if
server.ScriptTimeout=99999
Submit=request("submit")

if Submit<>"" Then
If Request("MailTitle") = "" Then  
response.write "<script LANGUAGE='javascript'>alert('��ܰС��ʾ���ʼ����ⲻ��Ϊ��!');history.go(-1);</script>"
Response.End
End If
If mailbody = "" Then  
response.write "<script LANGUAGE='javascript'>alert('��ܰС��ʾ���ʼ����ݲ���Ϊ��!');history.go(-1);</script>"
Response.End
End If

	MailTitle=Request("MailTitle")	
	UseHtml=Request("UseHtml")		
	Target=Request("Target")

	if Target="ALL" then'ȫ����Ա		
        emails=AllEmail("")
	elseif Target="ordinary" then'��ͨ��Ա
		emails=ordinaryEmail("")
	elseif Target="VIP" then'VIP��Ա
		emails=vipEmail("")
	elseif Target="store" then'���̻�Ա
		emails=storeEmail("")
	elseif Target="maillist" then'������־��Ա
		emails=maillistEmail("")
	elseif Target="custom" then'�Զ�������
		emails=Request("usermails")
	end if

	if Target<>"custom" then'��������Զ��尴|����
	SplitEmail=split(emails,"|")
	else
	SplitEmail=split(emails,chr(13))'�Զ��尴�ո�����
	end if
	if UseHtml="" then
		CanHTML=false
	else
		CanHTML=true
	end If	
    C_M_Chk=True'���ͣ�����ֵ�����ã�Smtp�������Ƿ���Ҫ�����֤ ,True��Ҫ��֤��False����Ҫ��֤��
	for i=0 to ubound(SplitEmail)
		tomail=trim(trim(SplitEmail(i)))
		if instr(tomail,"@")>0 And ChkMail(tomail)=True then			
			Call DNJYEmail(tomail,webname,MailTitle,mailbody,CanHTML,C_M_Chk)'�����ʼ����������ʼ�
			response.write "."
		end If		
	Next
	response.write "<script LANGUAGE='javascript'>alert('���ͳɹ���һ��"&ubound(SplitEmail)+1&"���ʼ�!');history.go(-1);</script>"	
end if

'��ȡ�����ַ��������������������������������������������������������������������
function AllEmail(userid)'��ȡȫ����Ա����
    set rs=server.createobject("adodb.recordset")
	sql="select email from DNJY_user where useryz=1"
	rs.open sql,conn,1,1	
	if not rs.eof Then
		do while not rs.eof
			thisEmail=rs(0)
			if instr(sssstemp,thisEmail)=0 then
			sssstemp=sssstemp&thisEmail&"|"
			end if
		rs.movenext
		loop
	end if
	rs.close
	AllEmail=sssstemp
end Function

function ordinaryEmail(userid)'��ȡ��ͨ��Ա����
    set rs=server.createobject("adodb.recordset")
	sql="select email from DNJY_user where useryz=1 and vip=0"
	rs.open sql,conn,1,1	
	if not rs.eof Then
		do while not rs.eof
			thisEmail=rs(0)
			if instr(sssstemp,thisEmail)=0 then
			sssstemp=sssstemp&thisEmail&"|"
			end if
		rs.movenext
		loop
	end if
	rs.close
	ordinaryEmail=sssstemp
end Function

function vipEmail(userid)'��ȡVIP��Ա����
    set rs=server.createobject("adodb.recordset")
	sql="select email from DNJY_user where useryz=1 and vip=1"
	rs.open sql,conn,1,1	
	if not rs.eof Then
		do while not rs.eof
			thisEmail=rs(0)
			if instr(sssstemp,thisEmail)=0 then
			sssstemp=sssstemp&thisEmail&"|"
			end if
		rs.movenext
		loop
	end if
	rs.close
	vipEmail=sssstemp
end Function

function storeEmail(userid)'��ȡ���̻�Ա����
    set rs=server.createobject("adodb.recordset")
	sql="select email from DNJY_user where useryz=1 and sdname<>''"
	rs.open sql,conn,1,1	
	if not rs.eof Then
		do while not rs.eof
			thisEmail=rs(0)
			if instr(sssstemp,thisEmail)=0 then
			sssstemp=sssstemp&thisEmail&"|"
			end if
		rs.movenext
		loop
	end if
	rs.close
	storeEmail=sssstemp
end Function

function maillistEmail(userid)'��ȡ������־��Ա����
    set rs=server.createobject("adodb.recordset")
	sql="select email from DNJY_user where useryz=1 and maillist=1"
	rs.open sql,conn,1,1	
	if not rs.eof Then
		do while not rs.eof
			thisEmail=rs(0)
			if instr(sssstemp,thisEmail)=0 then
			sssstemp=sssstemp&thisEmail&"|"
			end if
		rs.movenext
		loop
	end if
	rs.close
	maillistEmail=sssstemp
end Function

'��ȡ�����ַEND��������������������������������������������������������������������%>
<html>
<head>
<title>�ʼ�Ⱥ��</title>
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
    if((!checkemail(form1.usermails.value))&&(document.form1.usermails.value!=""))
	{
    alert("������Email��ַ����ȷ������������!");
    document.form1.usermails.focus();
    return false;
    }
if (document.form1.MailTitle.value=="" )
	{
	  alert("�����ʼ�����")
	  document.form1.MailTitle.focus()
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
<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1'>
<tr><td colspan=4 bgcolor="#BDBEDE" height=28 align=center><B>Ⱥ���ʼ�</b></td></tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="0" class="border">
 <form name="form1" method="post" action="" onSubmit="return formcheck()">
 <tr> <td width="19%" align="center" colspan="2" class="tdbg"><font color="#ff0000"> ע�⣺���ںܶ��ʾ֡��������������Ⱥ���ʼ����ܼ�⣬��������ܻ�������ͬ������ʱ�����ý�ֹ��������ķ��Ź��ܣ�����辭��Ⱥ���ʼ��Ľ��鹺�������Ƶ���ҵ�ʾ֣�</font></td></tr> 
<tr> <td width="19%" align="right" class="tdbg"><b>Ŀ&nbsp;&nbsp;&nbsp;&nbsp;�꣺</b> </td><td width="81%" class="tdbg"> <input name="Target" type="radio" value="ALL" checked> 
ȫ����Ա <input type="radio" name="Target" value="ordinary"> ��ͨ��Ա <input type="radio" name="Target" value="VIP"> VIP��Ա <input type="radio" name="Target" value="store">���̻�Ա <input type="radio" name="Target" value="maillist"> ������־��Ա <input type="radio" name="Target" value="custom"> 
�����Զ����ռ�������</td></tr> 
<tr> <td width="19%" align="right" class="tdbg"><b>�Զ����ռ������䣺</b></td><td width="100%" class="tdbg"><textarea name="usermails" cols="84" rows="9" wrap="VIRTUAL" id="usermails"></textarea>
    ��<font color="#ff0000">����ÿ��һ�����س���ӣ����򲻶��޷����ͣ�</font>��<p></td></tr>

<tr> 
<td width="19%" align="right" class="tdbg"><b>�ʼ���ʽ��</b></td><td width="81%" class="tdbg"><input type="radio" name="UseHtml" value="text" checked  style="border:1px black solid">                                                                   
      �ı��ʼ�&nbsp; <input type="radio" name="UseHtml" value="HTML" style="border:1px black solid">                                                                          
        HTML�ʼ�&nbsp;<p></td></tr> 
<tr> <td width="19%" align="right" class="tdbg"><b>�ʼ����⣺</b></td><td width="81%" class="tdbg"> 
<input name="MailTitle" type="text" id="MailTitle" size="60"> <font color="#ff0000"> *</font><p></td></tr> 
<tr> <td width="19%" align="right" class="tdbg"></td><td width="81%" class="tdbg"><font color="#ff0000"> �����������ݿ�Ҫ����ѡ����ʼ���ʽ��Ӧ��д������ͬʱ��д��</font> </td></tr> 
<tr> <td width="19%" align="right" valign="top" class="tdbg"><b>�ı���ʽ�ʼ����ݣ�</b></td><td width="81%" class="tdbg"> 
<textarea name="mailbody" cols="84" rows="15" wrap="VIRTUAL" id="mailbody"></textarea> 
 <font color="#ff0000"> *</font></td></tr>
<tr> <td width="19%" align="right" valign="top" class="tdbg"><b>HTML��ʽ�ʼ����ݣ�</b></td><td width="81%" class="tdbg"> 
<textarea id="editor" name="mailbodyhtm"  style="width:550px;height:280px;visibility:visible"></textarea>
 <font color="#ff0000"> *</font></td></tr>
<tr> <td width="19%" height="36" align="right" class="tdbg">&nbsp;</td><td width="81%" height="36" class="tdbg"> 
<input type="submit" name="Submit" value="ȷ���������ʼ�  "> <input type="reset" name="Submit2" value="  ȫ����д  "> 
</td></tr>

<tr>
  <td align="right" class="tdbg">&nbsp;</td>
  <td class="tdbg">&nbsp;</td>
</tr>
<tr>
  <td align="right" class="tdbg">&nbsp;</td>
  <td class="tdbg">&nbsp;</td>
</tr>
</form></table>
</body>
</html>
<!--#include virtual="/Include/bottom_superadmin.asp" -->
