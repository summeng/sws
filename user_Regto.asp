<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=top.asp-->
<!--#include file="Include/err.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%if not ChkPost() then 
response.write "��ֹվ���ύ�����"
response.end 
end If
'if request.cookies("reg")("regkey")="" Or Instr(Request.ServerVariables("HTTP_REFERER"),""&reg&"")=0 then 
'response.redirect ""&reg&"?"&C_ID&"" 
'end If

randomize
session("sms_code")= CInt(8999 * Rnd + 1000)'�����ֻ���֤����� 

if lmkg2="1" then
call errdick(50)
response.end
end If

Dim tjname
tjname=trim(request("tjname"))%>
<link rel="stylesheet" href="/<%=strInstallDir%>css/css.css" type="text/css">
<style type="text/css">
/*��ʾ��Ϣ*/
.info {position:relative;background:#;color:#666; text-decoration:none;font-size:12px;width:0px;height:0px;text-align:center;border:0px solid #ccc;line-height:18px;}/*�������ӵ�����,һ��Ҫ����Ϊrelative����ʹ��ʾ�����������*/
.info:hover {background:#eee;color:#333;}
.info span {display: none }/*���������µ�spanΪ����״̬*/
.info:hover span /*����hover�µ�span����Ϊ����״̬,��������ʾ���λ��*/{display:block;position:absolute;top:0px;left:25px;width:200px;
border:1px solid #DC5826; background:#FFF3B4; color:#333;padding:5px;text-align:left;}
</style>
<script language="javascript" type="text/javascript">
//����ע���û����Ƿ��ظ��ж�
var xmlHttp = false;
try {
  xmlHttp = new ActiveXObject("Msxml2.XMLHTTP");
} catch (e) {
  try {
    xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
  } catch (e2) {
    xmlHttp = false;
  }
}
if (!xmlHttp && typeof XMLHttpRequest != 'undefined') {
  xmlHttp = new XMLHttpRequest();
}

function callServer() {
//ֻ������һ��document.getElementById("username").value;�е�usernameΪ���е��û�������
  var username = document.getElementById("username").value;
  if ((username == null) || (username == "")) return;

var reId=/^[\w\u0391]+$/;//ĿǰΪ֧�ִ�Сд��ĸ�����ֺ��»����ַ��������\uFFE5Ϊ���ֿ���ʹ�ã���var reId=/^[\w\u0391-\uFFE5]+$/;
var b_id=reId.test(username);
if(!b_id){
var url = "user_regxml.asp?name=" + escape("yes");
xmlHttp.open("GET", url, true);
xmlHttp.onreadystatechange = updatePage
xmlHttp.send(null);
}
	else
  var url = "user_regxml.asp?name=" + escape(username);
  xmlHttp.open("GET", url, true);
  xmlHttp.onreadystatechange = updatePage;
  xmlHttp.send(null);  
}

function updatePage() {
  if (xmlHttp.readyState < 4) {
	test.innerHTML="loading...";
  }
  if (xmlHttp.readyState == 4) {
    var response = xmlHttp.responseText;
	test.innerHTML=response;
  }

}


</script>
<script language="javascript">
<!--
//�ֽ������ƣ����������=========
String.prototype.Length=function(){
temp=this.replace(/([^\x00-\xff])/g,"$1$1");
return temp.length
}
//��ʾ�����ַ���

function kn()
  {
   var ln=thisForm.username.value.Length();
     num.innerText='������������:'+(16-ln)+'���ַ�';
      //if (ln>=16) form.username.readOnly=true;  //��������������������޷�������޸�
}
function kn2()
 {
   var ln=thisForm.password.value.Length();
     num2.innerText='������������:'+(12-ln)+'���ַ�';
      
}
// --></script>
<script language="javascript" type="text/javascript">
<!--
//��֤��Ա�ʺ�����ȷ��
function checkeid(username){
var str=username;
//��JavaScript�У�������ʽֻ��ʹ��"/"��ͷ�ͽ���������ʹ��˫����
var Expression=/^[\w\u0391]+$/; //֧�ִ�Сд��ĸ�����ֺ��»����ַ�
var objExp=new RegExp(Expression);
if(objExp.test(str)==true){
return true;
}else{
return false;
}
} 

//��֤���֤��ȷ��
function checkeNO(NO){
var str=NO;
//��JavaScript�У�������ʽֻ��ʹ��"/"��ͷ�ͽ���������ʹ��˫����
var Expression=/\d{17}[\d|X]|\d{15}/; 
var objExp=new RegExp(Expression);
if(objExp.test(str)==true){
return true;
}else{
return false;
}
} 


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


function CheckForm()
{

        if(document.thisForm.username.value.length<1)
	{
	    alert("��¼�ʺŲ���Ϊ��!");
		document.thisForm.username.focus()
	    return false;
	}
     if ((document.thisForm.username.value.Length()>16) || (document.thisForm.username.value.Length()<1)) //�ֽ������ƣ����������
     {
	 alert("��¼�ʺų���Ҫ��1��16�ֽڣ����������룡");
	  document.thisForm.username.focus()
	  return false
  }
    if (!checkeid(thisForm.username.value))
	{
    alert("��¼�ʺŽ�֧�ִ�Сд��ĸ�����ֺ��»����ַ�������������!");
    document.thisForm.username.focus();
    return false;
    }

	  if ((document.thisForm.password.value.Length()>12) || (document.thisForm.password.value.Length()<5)) //�ֽ������ƣ����������
     {
	 alert("���볤��Ҫ��5��12�ֽڣ����������룡");
	  document.thisForm.password.focus()
	  return false
  }
	    if(document.thisForm.password1.value.length<1)
	{
	    alert("ȷ�����벻��Ϊ��!");
		document.thisForm.password1.focus()
	    return false;
	}
	   if(document.thisForm.password1.value!=document.thisForm.password.value)
        {
       alert("�����������벻һ��!");
	   document.thisForm.password1.focus()
       return false;
        }
    if (document.thisForm.mobile.value=="" )
	{	  
      alert("�ֻ����벻��Ϊ�գ����������룡");
	  document.thisForm.mobile.focus()
	  return false
	 }
	var sMobile = document.thisForm.mobile.value
	if((document.thisForm.mobile.value!="") && (!(/^((\(\d{3}\))|(\d{3}\-))?13[0-9]\d{8}|14[0-9]\d{8}|15[0-9]\d{8}|17[0-9]\d{8}|18[0-9]\d{8}/.test(sMobile)))) //(�������15,18,17�����Ŷ�)
	{
		alert("����������11λ�ֻ��Ż�����ȷ���ֻ���ǰ��λ");
		document.thisForm.mobile.focus();
		return false;
	}
 <%if sms_kg=True then%>
	
    if (document.thisForm.captcha.value=="" )
	{	  
      alert("�ֻ�����У���벻��Ϊ�գ����������룡");
	  document.thisForm.captcha.focus()
	  return false
	 }
    if(document.thisForm.captcha.value != <%=session("sms_code")%>) 
	{	  
      alert("�ֻ�����У������ϵͳ���ɵĲ�������ʱ��̫��ʧЧ����ϵ�ͷ�ɾ����¼�����»�ȡ�ֻ�����У���룡");
	  document.thisForm.captcha.focus()
	  return false
	 }
 <%end if%>
<%If mailqr=1 then%>
   if(document.thisForm.email.value.length<1)
	{
	    alert("�����ַ����Ϊ��!");
		document.thisForm.email.focus()
	    return false;
	}
 <%end if%>
    if ((document.thisForm.email.value!="")&&(!checkemail(thisForm.email.value)))
	{
    alert("������Email��ַ����ȷ������������!");
    document.thisForm.email.focus();
    return false;
    }
   if(document.thisForm.email.value.length>30 )
	{
	    alert("���䲻�ܳ���30���ַ�!");
	    document.thisForm.email.focus();
	    return false;
	}

 <%if codekg1=1 then%>
    if (document.thisForm.wenti.value=="" )
	{	  
      alert("��֤�𰸲���Ϊ�գ����������룡");
	  document.thisForm.wenti.focus()
	  return false
	 }
	
    if(document.thisForm.wenti.value != document.thisForm.daan.value) 
	{	  
      alert("��֤�𰸲��ԣ����������룡");
	  document.thisForm.wenti.focus()
	  return false
	 }
 <%end if%>
 <%if codekg2=1 then%>
    if (document.thisForm.yzm.value=="" )
	{	  
      alert("������֤�벻��Ϊ�գ����������룡");
	  document.thisForm.yzm.focus()
	  return false
	 }
<%end if%>
 <%if codekg3=1 then%>
    if (document.thisForm.code.value=="" )
	{	  
      alert("������֤�벻��Ϊ�գ����������룡");
	  document.thisForm.code.focus()
	  return false
	 }
<%end if%>
<%if codekg4=1 then%>
    if (document.thisForm.ttdv.value=="" )
	{	  
      alert("��֤���ڲ���Ϊ�գ����������룡");
	  document.thisForm.ttdv.focus()
	  return false
	 }
<%end if%>
<%if codekg5=1 then%>
    if (document.thisForm.validate_codee.value=="" )
	{	  
      alert("��ʽ��֤�벻��Ϊ�գ����������룡");
	  document.thisForm.validate_codee.focus()
	  return false
	 }
<%end if%>
	if((document.thisForm.answer1.value=="")&&(document.thisForm.problem.value!=""))
        {
            alert("����ȡ�����������Ҫ���!");
			document.thisForm.answer1.focus()
            return false;
        }
    if((document.thisForm.problem.value.length>=1) && (document.thisForm.problem.value == document.thisForm.answer1.value)) {
	document.thisForm.answer1.focus();
	document.thisForm.answer1.value = '';
	document.thisForm.answer2.value = '';
    alert("����������𰸲�Ҫ��ͬ�����������룡");
	return false;
  }

	   if(document.thisForm.answer1.value!=document.thisForm.answer2.value)
        {
            alert("��������𰸲�һ��!");
			document.thisForm.answer2.focus()
            return false;
        }
    if ((!checkeNO(thisForm.idcard.value)) && (document.thisForm.idcard.value!=""))
   {
    alert("���������֤���벻��ȷ!");
    document.thisForm.idcard.focus();
    return false;
    }
//var sMobile = document.thisForm.dianhua.value
//if((document.thisForm.dianhua.value!="") && (!(/^([0\+]\d{2,3}-)?(0[0-9]{2,3}\-)?([2-9][0-9]{6,7})+(\-[0-9]{1,4})?$/.test(sMobile)))) //(���Ҵ���-����-�绰-�ֻ���ʽ(��ѡ)��+086-010-85858585-2538)
//	{
//		alert("���淶�ĵ绰����");
//		document.thisForm.dianhua.focus();
//		return false;
//	}
if(document.thisForm.dianhua.value.length>30 )
	{
	    alert("�̶��绰���ܳ���30���ַ�!");
		document.thisForm.dianhua.focus();
	    return false;
	}

//var sMobile = document.thisForm.fax.value
//if((document.thisForm.fax.value!="") && (!(/^([0\+]\d{2,3}-)?(0[0-9]{2,3}\-)?([2-9][0-9]{6,7})+(\-[0-9]{1,4})?$/.test(sMobile)))) //(���Ҵ���-����-�绰-�ֻ���ʽ(��ѡ)��+086-010-85858585-2538)
//	{
//		alert("���淶�Ĵ���绰����");
//		document.thisForm.fax.focus();
//		return false;
//	}
if(document.thisForm.fax.value.length>30 )
	{
	    alert("���治�ܳ���30���ַ�!");
		document.thisForm.fax.focus();
	    return false;
	}

 if (document.thisForm.http.value=="http://" )     
  {   
  alert("��ַǰ�治Ҫ��http://")
  document.thisForm.http.focus();
  return false;  
  }   
  
}
//-->
</SCRIPT>
<!--������֤����ÿ�ʼ-->
<SCRIPT LANGUAGE=javascript>
/*��ʾ��֤�� o start1*/
function get_Code() {
        var Dv_CodeFile = "Dv_GetCode.asp";
        if(document.getElementById("imgid"))
                document.getElementById("imgid").innerHTML = '<img src="'+Dv_CodeFile+'?t='+Math.random()+'" alt="���ˢ����֤��" style="cursor:pointer;border:0;vertical-align:middle;height:30px;" onclick="this.src=\''+Dv_CodeFile+'?t=\'+Math.random()" />'
}
/*o end*/
</script>
<script language="JavaScript" type="text/javascript">
var dvajax_request_type = "GET";
</script>
<script language="JavaScript" src="dv_ajax.js" type="text/javascript"></script>
<!---������֤����ý���-->
<!---�ֻ�����У������ÿ�ʼ-->
<script language="javascript">
	function get_mobile_code(){
//�ж��Ƿ������ֻ���
        var myReg = /^1[3,4,5,7,8]\d{9}$/g;
        var mobile = document.thisForm.mobile.value;
        if (!myReg.exec(mobile)) {
            alert("������11λ��ȷ���ֻ����룡" + mobile);
            document.thisForm.mobile.focus();
            return false;
        }
//����У����
        $.post('sms.asp', {mobile:jQuery.trim($('#mobile').val())}, function(msg) {
            alert(jQuery.trim(unescape(msg)));
        });
	};
</script>
<script type="text/javascript" src="Include/sms_jQuery-1.6.js"></script>
<script type="text/javascript" src="Include/sms_button.js"></script>
<!---�ֻ�������֤���ý���-->
<script language="javascript" type="text/javascript">
<!--
	function checkPwd()
	{
	}
	first("selectp","selectc","thisForm",0,0);
	function showadv()
	{
		if (document.thisForm.advshow.checked == true) 
		{
			document.getElementById("adv").style.display = "inline";
			document.getElementById("advance").innerHTML="�رո߼�����ѡ��";
		}
		else
		{
			document.getElementById("adv").style.display = "none";
			document.getElementById("advance").innerHTML="��ʾ�߼�����ѡ��";
		}
	}
//-->
</script> 
</HEAD>
<center>
<TABLE width=980 border=0 cellpadding=3 cellspacing=1 align=center  bgcolor=#FFFFFF    class="grayline"  background="/img/hd_normal_bg3.gif">
<tr><td height=30>Ŀǰλ�ã�<a href=index.asp?<%=C_ID%>>��ҳ</a> > ��Աע��ڶ���</td></tr>
</TABLE>
<TABLE width=980 border=0 cellpadding=3 cellspacing=1 align=center  bgcolor=#FFFFFF height=550    class="grayline" >
<tr><td height="1" background="images/bgline.gif"></td></tr>
<tr><td>
<body topmargin="0" leftmargin="0">
<div>
  <center>
  <table width="980" height="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#111111" bgcolor="#B6E688" style="border-collapse: collapse">
    <tr>
      <td height="156"><div >
        
        <br><table width="800" align=center border="0" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">    
                <tr>                 
                  <td width="800" bgcolor="#FFFFFF">
                    <table width="800" height="126" border="1" cellpadding="3" cellspacing="3" bordercolor="#CCCCCC">
					
                   <form name="thisForm" method="post" action="user_reg_save.asp?<%=C_ID%>" language="javascript" onsubmit="return CheckForm()">   
                      <tr bgcolor="#FFFFFF"> 
                        <td width="654" height="21" colspan=2>
                        <p align="CENTER"><strong>Ϊȷ����Ϣ��ʵ��Ч����������дע����Ϣ (��<font color="#ff0000">*</font>����)</strong></td>
                      </tr>
		             <tr bgcolor="#6CCEE5"> 
                        <td height="30" align="center" colspan="2">
                        <font color="#ff0000">�� �� ��</font></td>
                    </tr>
                   <tr bgcolor="#FFFFFF"> 
                      <td height="30">
                      <p align="right">��½�ʺţ�</td>
                      <td height="30">
                      &nbsp;
                      <input class="inputa" type="text" id="username" maxlength="16" name="username" size="20" onpropertychange="kn()" autocomplete="off" onChange="callServer();"><font color="#ff0000"> *</font> &nbsp;<span id="test"><img src='images/reg_warning.gif'>�������ʺż����Ч��</span><br>&nbsp;&nbsp;<span id=num>�㻹��������16���ַ�</span>&nbsp;&nbsp;(������Ӣ�ĺ�����)</td>
                    </tr>
 
                      </tr>
                   <tr bgcolor="#FFFFFF"> 
                      <td height="30">
                      <p align="right">��¼���룺</td>
                      <td height="30">
                      &nbsp;
                        <input class="inputa" type="password" name="password" size="20" onpropertychange="kn2()"><font color="#ff0000"> *</font> <span id=num2>5-12λ,�㻹��������12���ַ�</span></td>
                      </tr>
                      </tr>
                   <tr bgcolor="#FFFFFF"> 
                      <td height="30">
                      <p align="right">ȷ�����룺</td>
                      <td height="30">
                      &nbsp;
                        <input class="inputa" type="password" name="password1" size="20" ><font color="#ff0000"> *</font> &nbsp;&nbsp;&nbsp;</td>
                      </tr>
                      </tr>

 
                    <tr bgcolor="#FFFFFF"> 
                      <td height="30">
                      <p align="right">��&nbsp;&nbsp;&nbsp;&nbsp;��</td>
                      <td height="30">
                      &nbsp;
                      <select id="ctlSex" name="Sex" class="inputa">
                        <option value="1" selected>��</option>
                        <option value="0">Ů</option>
                      </select>  <font color="#ff0000"> *</font></td>
                    </tr>

                     <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      �ֻ����룺</td>
                      <td height="30">
                      &nbsp;
                      <input class="inputa" type="text" maxlength="30" id="mobile" name="mobile" size="24" onKeyUp="if(isNaN(value)){alert('�ƶ��绰ֻ������������');value='';}"> <font color="#ff0000"> *</font> </td>
                    </tr>
					<%if sms_kg=True then%>
                     <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      У �� �룺</td>
                      <td height="30">
                      &nbsp;
                      <input class="inputa" type="text" size="8" name="captcha" class="inputBg" /> <font color="#ff0000"> *</font> �鿴�ֻ�����У���벢���ڴ� <input id="zphone" type="button" value=" ��ȡ�ֻ�У���� " onclick="get_mobile_code();countDown(this,30);" ></td>
                    </tr>  
					<%End if%>
          			<tr bgcolor="#FFFFFF"> 
                        <td height="30" align="right">
                        �������䣺</td>
                        <td height="13">
                      &nbsp;
                      <input class="inputa" type="text" maxlength="255" name="email" size="24" > <%If mailqr=1 then%><font color="#ff0000"> *</font>  ע�᱾վ��Ա���ʼ���֤��<%End if%>������ȷ���䣡</td>
                    </tr>
                   <tr bgcolor="#FFFFFF"> 
                      <td height="30">
                      <p align="right">�� �� �ˣ�</td>
                      <td height="30">
                      &nbsp;
                      <input class="inputa" type="text" maxlength="16" name="recommend" value="<%=tjname%>" size="24"> ��д�Ƽ��˻�ԱID�ſɸ������ӻ��� <%=jf_5%> �֣�û�������ա�</td>
                    </tr>

<%Call OpenConn
if codekg1=1 then
	  '�ʴ�ʽ��֤
dim rscheck
Randomize 
Set rscheck= Server.CreateObject("ADODB.Recordset")
If SystemDatabaseType = 1 Then
sql="select  * from DNJY_wenda where xs=1 order BY newid()"
Else
sql="select  * from DNJY_wenda where xs=1 order BY rnd(-(ID+"& rnd() &"))"
End if
rscheck.open sql,conn,1,1
%>
                     <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      �ʴ���֤��</td>
                      <td height="30">
                      &nbsp;���⣺<%=rscheck("wenti")%><br>&nbsp;�𰸣�<input type="text" name="wenti" size="12" class="inputa"> <font color="#ff0000">*</font></td>
                    </tr><input type="hidden" name="daan" value="<%=rscheck("daan")%>">
	<%rscheck.close
	set rscheck=Nothing
	End If
	if codekg2=1  then%>
                     <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      ������֤��</td>
                      <td height="30">
                      &nbsp;<input type="text" name="yzm" size="4" class="inputa"/><img src="tt_getcode.asp" height="18" alt="��֤��,�������?����ˢ����֤��" style="cursor : pointer;" onclick="this.src='tt_getcode.asp'"/>&nbsp;<font color="#ff0000">*</font>&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>��֤��,�������?����ˢ����֤��</span></a></td>
                    </tr>
	  <%End if
	if codekg3=1 then%>
                     <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      ������֤��</td>
                      <td height="30">
                      &nbsp;<input type="text" class="inputa" name="code" id="code" size="4" maxlength="4" tabindex="6" onfocus="get_Code();this.onfocus=null;" onkeyup="dv_ajaxcheck('checke_dvcode','code');"  autocomplete="off"/>
    <font color="#ff0000">*</font> <span id="imgid" style="color:red">�����ȡ��֤��</span><span id="isok_code"></span></td>
                    </tr>
<%End if%>
<%if codekg4=1 then%>
                     <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      ������֤��</td>
                      <td height="30">
                      &nbsp������ <select name="ttdv" class="inputa">
<option value="" selected="selected">��ѡ��</option>
<option value="1">����һ</option>
<option value="2">���ڶ�</option>
<option value="3">������</option>
<option value="4">������</option>
<option value="5">������</option>
<option value="6">������</option>
<option value="0">������</option>
</select>&nbsp;&nbsp; <font color="#ff0000">*</font><a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>��ѡ����ȷ�����ڼ�</span></a></td>
      </tr>
<%End if%>
<%if codekg5=1 then%>
                     <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      ��ʽ��֤��</td>
                      <td height="30">
                      &nbsp;<img src="tt_getcodee.asp" alt="��֤��,�������?����ˢ����֤��" height="10" style="cursor : pointer;" onClick="this.src='tt_getcodee.asp?t='+(new Date().getTime());" /><input name="validate_codee" class="inputa" type="text" tabindex="3" value="" size="4" maxlength="4" onkeyup="value=value.replace(/[^\d|^\.]/g,'')"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))"/> &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>��֤��,�������?����ˢ����֤��</span></a></td>
                    </tr>
<%End if%>

 			<tr bgcolor="#6CCEE5"> 
                        <td height="30" align="center" colspan="2">
                        <font color="#ff0000"><INPUT id=advcheck name=advshow type=checkbox value=1 onclick=showadv()>
         <span id=advance>��ʾ�߼�����ѡ��</span>�����������������Ϸ����Ժ󷢲���Ϣ��</font></td>
                    </tr>
</table>
<table style="DISPLAY: none" id=adv width="800" height="126" border="1" cellpadding="3" cellspacing="3" bordercolor="#CCCCCC">
                  <tr bgcolor="#FFFFFF"> 
                      <td height="30" width="88">
                      <p align="right">�������⣺</td>
                      <td height="30">
                      &nbsp;
                        <input class="inputa" type="problem" maxlength="20" name="problem" size="20" > ��д�󷽱����������һ�</td>
                      </tr>
                    
                   <tr bgcolor="#FFFFFF"> 
                      <td height="30">
                      <p align="right">����𰸣�</td>
                      <td height="30">
                      &nbsp;
                        <input class="inputa" type="answer1" maxlength="20" name="answer1" size="20" >���һ�����ʱҪ�ش�</td>
                      </tr>                     
                      
                   <tr bgcolor="#FFFFFF"> 
                      <td height="30">
                      <p align="right">ȷ�ϴ𰸣�</td>
                      <td height="30">
                      &nbsp;
                        <input class="inputa" type="answer2" maxlength="20" name="answer2" size="20" > &nbsp;&nbsp;&nbsp;</td>
                      </tr>
                   <tr bgcolor="#FFFFFF"> 
                      <td height="30">
                      <p align="right">��ʵ������</td>
                      <td height="30">
                      &nbsp;
                      <input class="inputa" type="text" maxlength="12" name="name" size="20"> </td>
                    </tr>

                    <tr bgcolor="#FFFFFF"> 
                      <td height="30">
                      <p align="right">�� �� ֤��</td>
                      <td height="30">
                      &nbsp;
                      <input class="inputa" type="text" maxlength="20" name="idcard" size="24" > �������⹫����</td>
                    </tr>         

       
			<tr bgcolor="#FFFFFF">
			 <td height="30" align="right">QQ&nbsp;���룺</td>
			<td height="14">��<input class="inputa" type="text" maxlength="255" name="qq" size="24" onKeyUp="if(isNaN(value)){alert('QQ����ֻ������������');value='';}"></td>
			        </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      ΢�ź��룺</td>
                      <td height="30">
                      &nbsp;
                      <input class="inputa" type="text" maxlength="50" name="weixin" size="24" ></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      ��˾���ƣ�</td>
                      <td height="30">
                      &nbsp;
                      <input class="inputa" type="text" maxlength="80" name="mpname" size="37" > Ҳ����ҵ��Ƭ����</td>
                    </tr>					  
                    <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      ͨ�ŵ�ַ��</td>
                      <td height="30">
                      &nbsp;
                      <input class="inputa" type="text" maxlength="100" name="dizhi" size="60" ></td>
                    </tr>
                     <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      �������룺</td>
                      <td height="30">
                      &nbsp;
                      <input class="inputa" type="text" maxlength="6" name="youbian" size="12"  onKeyUp="if(isNaN(value)){alert('��������ֻ������������');value='';}">    </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      �̶��绰��</td>
                      <td height="30">
                      &nbsp;
                      <input class="inputa" type="text" maxlength="30" name="dianhua" size="37" > ���Ҵ���-����-�绰-�ֻ���ʽ(��ѡ����)��+086-010-85858585-2538</td>
                    </tr>
                     <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      ��&nbsp;&nbsp;&nbsp;&nbsp;�棺</td>
                      <td height="30">
                      &nbsp;
                      <input class="inputa" type="text" maxlength="30" name="fax" size="37" > ���Ҵ���-����-�绰-�ֻ���ʽ(��ѡ����)��+086-010-85858585-2538</td>
                    </tr>
                     <tr bgcolor="#FFFFFF"> 
                      <td height="30" align="right">
                      ��&nbsp;&nbsp;&nbsp;&nbsp;ַ��</td>
                      <td height="30">
                      &nbsp;  
                      <input class="inputa" type="text" value=""  maxlength="100" name="http" size="60"> ����www.ip126.com��ǰ�治�ܴ�http:// </td>
                    </tr>
			<tr bgcolor="#FFFFFF"> 
                        <td height="30" align="right">
                        ����������</td>
                        <td height="13">
                      &nbsp;
                      <%
set rs=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")
%>
          <SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
	dim count5:count5 = 0
        do while not rs.eof 
        %>
subcat[<%=count5%>] = new Array("<%=rs("city")%>","<%=rs("id")%>","<%=rs("twoid")%>");
        <%
        count5 = count5 + 1
        rs.movenext
        loop
Call CloseRs(rs)
        %>
onecount=<%=count5%>;
          </SCRIPT>
          <%
set rs=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid>0")
%>
          <SCRIPT language = "JavaScript">
var onecount4;
onecount4=0;
subcat4 = new Array();
        <%
		dim count4:count4 = 0
        do while not rs.eof 
        %>
subcat4[<%=count4%>] = new Array("<%=rs("city")%>","<%=rs("id")%>","<%=rs("twoid")%>","<%=rs("threeid")%>");
        <%
        count4 = count4 + 1
        rs.movenext
        loop
Call CloseRs(rs)
        %>
onecount4=<%=count4%>;

function changelocation(locationid)
    {
    document.thisForm.city_two.length = 0; 
	document.thisForm.city_two.options[0] = new Option('ѡ�����','');
	document.thisForm.city_three.length = 0; 
	document.thisForm.city_three.options[0] = new Option('ѡ�����','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.thisForm.city_two.options[document.thisForm.city_two.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
	
	function changelocation4(locationid,locationid1)
    {
    document.thisForm.city_three.length = 0; 
    document.thisForm.city_three.options[0] = new Option('ѡ�����','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount4; i++)
        {
            if (subcat4[i][2] == locationid)
            { 
			if (subcat4[i][1] == locationid1)
			{
                document.thisForm.city_three.options[document.thisForm.city_three.length] = new Option(subcat4[i][0], subcat4[i][3]);
            }        
        }
       } 
    }            
              </SCRIPT>
          <SELECT name="city_one" size="1" id="select" class="inputa" onChange="changelocation(document.thisForm.city_one.options[document.thisForm.city_one.selectedIndex].value)">
            <OPTION value="" selected>ѡ�����</OPTION>
            <%set rs=conn.execute("select * from DNJY_city where id>0 and twoid=0 order by indexid asc")
if rs.eof or rs.bof then
response.write "<option value=''>û�з���</option>"
else
do until rs.eof
response.write "<option value='"&rs("id")&"'>"&rs("city")&"</option>"
    rs.movenext
    loop
end if
Call CloseRs(rs)
%>
          </SELECT>
          <SELECT name="city_two" id="select6" class="inputa" onChange="changelocation4(document.thisForm.city_two.options[document.thisForm.city_two.selectedIndex].value,document.thisForm.city_one.options[document.thisForm.city_one.selectedIndex].value)">
            <OPTION value="" selected>ѡ�����</OPTION>
          </SELECT>
          <SELECT name="city_three" id="select7" class="inputa">
            <OPTION value="" selected>ѡ�����</OPTION>
          </SELECT> ��Ĭ�����ڵ�������д�����Ժ󷢲���Ϣ�Զ���д��</td>
                    </tr>
			<tr bgcolor="#FFFFFF"> 
                        <td height="30" align="right">
                        ��Ϣ���ࣺ</td>
                        <td height="13">
                      &nbsp;
                      <%
set rs=conn.execute("select * from DNJY_type where id>0 and twoid>0 and threeid=0")
%>
          <SCRIPT language = "JavaScript">
var onecount6;
onecount6=0;
subcat7 = new Array();
        <%
	dim count6:count6 = 0
        do while not rs.eof 
        %>
subcat7[<%=count6%>] = new Array("<%=rs("name")%>","<%=rs("id")%>","<%=rs("twoid")%>");
        <%
        count6 = count6 + 1
        rs.movenext
        loop
Call CloseRs(rs)
        %>
onecount6=<%=count6%>;
          </SCRIPT>
          <%
set rs=conn.execute("select * from DNJY_type where id>0 and twoid>0 and threeid>0")
%>
          <SCRIPT language = "JavaScript">
var onecount7;
onecount7=0;
subcat5 = new Array();
        <%
		dim count7:count7 = 0
        do while not rs.eof 
        %>
subcat5[<%=count7%>] = new Array("<%=rs("name")%>","<%=rs("id")%>","<%=rs("twoid")%>","<%=rs("threeid")%>");
        <%
        count7 = count7 + 1
        rs.movenext
        loop
Call CloseRs(rs)
        %>
onecount7=<%=count7%>;

function changelocation6(locationid)
    {
    document.thisForm.type_two.length = 0; 
	document.thisForm.type_two.options[0] = new Option('ѡ�����','');
	document.thisForm.type_three.length = 0; 
	document.thisForm.type_three.options[0] = new Option('ѡ�����','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount7; i++)
        {
            if (subcat7[i][1] == locationid)
            { 
                document.thisForm.type_two.options[document.thisForm.type_two.length] = new Option(subcat7[i][0], subcat7[i][2]);
            }        
        }
        
    }    
	
	function changelocation7(locationid,locationid1)
    {
    document.thisForm.type_three.length = 0; 
    document.thisForm.type_three.options[0] = new Option('ѡ�����','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount7; i++)
        {
            if (subcat5[i][2] == locationid)
            { 
			if (subcat5[i][1] == locationid1)
			{
                document.thisForm.type_three.options[document.thisForm.type_three.length] = new Option(subcat5[i][0], subcat5[i][3]);
            }        
        }
       } 
    }            
              </SCRIPT>
          <SELECT name="type_one" size="1" id="select" class="inputa" onChange="changelocation6(document.thisForm.type_one.options[document.thisForm.type_one.selectedIndex].value)">
            <OPTION value="" selected>ѡ�����</OPTION>
            <%set rs=conn.execute("select * from DNJY_type where id>0 and twoid=0 order by indexid asc")
if rs.eof or rs.bof then
response.write "<option value=''>û�з���</option>"
else
do until rs.eof
response.write "<option value='"&rs("id")&"'>"&rs("name")&"</option>"
    rs.movenext
    loop
end if
Call CloseRs(rs)
%>
          </SELECT>
          <SELECT name="type_two" id="select6" class="inputa" onChange="changelocation7(document.thisForm.type_two.options[document.thisForm.type_two.selectedIndex].value,document.thisForm.type_one.options[document.thisForm.type_one.selectedIndex].value)">
            <OPTION value="" selected>ѡ�����</OPTION>
          </SELECT>
          <SELECT name="type_three" id="select7" class="inputa">
            <OPTION value="" selected>ѡ�����</OPTION>
          </SELECT> ��Ĭ����Ϣ���࣬��д�����Ժ󷢲���Ϣ�Զ���д��</td>
                    </tr>
			      <tr bgcolor="#FFFFFF"> 
                        <td height="30" align="right">
                        ������Ѷ��</td>
                        <td height="13"> <input type="radio" name="maillist" value="1" >                 
                  ���붩��&nbsp;&nbsp;&nbsp;&nbsp;                          
                  <input type="radio" name="maillist" value="0" checked>                 
                  ���붩��
                      &nbsp;�����ĺ��������ձ�վ���͵���Ѷ��
                      </td>
                    </tr>					
                    </table>
                 
	         <table width="800" align=center border="0" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">		           
                    <tr bgcolor="#FFFFFF"> 
                      <td height="25" colspan="3"><p align="CENTER"> 
					  <input type="submit" value=" ȷ �� " name="yes" onclick="javascript:return CheckForm();" language="javascript" id="yes" >
                      <input name="reset" type="reset" value=" �� д ">			 
                </form>
                      </td>
                    </tr>
               
              </table>
</td></tr>
</table> 
</center>
</BODY>
</HTML>

<%
If mailqr=1 then
Dim rsdq
set rsdq=server.createobject("adodb.recordset")
sql="delete from [DNJY_usertemp] where DateDiff("&DN_DatePart_D&",zcdata,"&DN_Now&")>"&regyxq&""'ɾ��3��δ�ʼ���֤����ʱע����Ϣ
rsdq.open sql,conn,1,3
End If
Response.Cookies("reg")("regkey")=""

If sms_del>0 then
set rsdq=server.createobject("adodb.recordset")
sql="delete from [DNJY_sms] where DateDiff("&DN_DatePart_D&",sms_addtime,"&DN_Now&")>"&sms_del&""'ɾ��ָ���������ֻ�������֤��¼
rsdq.open sql,conn,1,3
End If

Call CloseDB(conn)%>