<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%If (request("shows1")=1 Or request("shows2")=1 Or request("shows3")=1 Or request("shows4")=1 Or request("shows5")=1 Or request("shows6")=1 Or request("shows7")=1 Or request("shows8")=1 Or request("shows9")=1 Or request("shows10")=1) And request("HtmlHome")=1 Then
Response.Write ("<script language=javascript>alert('ѡ�񰴵�������վ����ʽ��ʾ�й�����ʱ�������ɾ�̬��ҳ!');history.go(-1);</script>")
Response.end
End If
If request("ok")=1 And request("leixing")="" Then
Response.Write ("<script language=javascript>alert('��Ϣ���Ͳ���Ϊ�գ���Ҫע���ʽ����!');history.go(-1);</script>")
Response.end
End If
If request("ok")=1 And (request("S10")="" or request("S11")="") Then
Response.Write ("<script language=javascript>alert('QQ�ͷ�������ǳƲ���Ϊ�գ���Ҫע���ʽ!');history.go(-1);</script>")
Response.end
End If
If request("mailqr")=1 And request("regmail")=0 Then
Response.Write ("<script language=javascript>alert('ѡ���ʼ�ȷ��ʱ����ͬʱѡ���ʼ�֪ͨ!');history.go(-1);</script>")
Response.end
End If
If request("ok")=1 And request("vipje")<1  Then
Response.Write ("<script language=javascript>alert('[VIP��Ա�ʸ���Ҫ]ѡ��Ϊ�������0������!');history.go(-1);</script>")
Response.end
End If


dim ObjInstalled
Const Script_FSO="Scripting.FileSystemObject"
ObjInstalled=IsObjInstalled(Script_FSO)
If ObjInstalled=false Then
 Response.Write "<b><font color=red>��ķ�������֧�� FSO(Scripting.FileSystemObject)! ����ʹ�ñ����ܡ� ��ֱ���޸ġ���Ŀ¼�µ�config.asp��Database.asp���ļ��е����ݡ�</font></b>"
End If
%>
<%call checkmanage("02")
Dim HomeElite,HomeElite1,HomeElite2,HomeElite3,HomeElite4,HomeElite5,s,url	
id=request("id")   
codekg1=request.form("codekg1")
codekg2=request.form("codekg2")
codekg3=request.form("codekg3")
codekg4=request.form("codekg4")
codekg5=request.form("codekg5")
If request.form("codekg1")="" Then codekg1="0"
If request.form("codekg2")="" Then codekg2="0"
If request.form("codekg3")="" Then codekg3="0"
If request.form("codekg4")="" Then codekg4="0"
If request.form("codekg5")="" Then codekg5="0"
codekg=codekg1+"|"+codekg2+"|"+codekg3+"|"+codekg4+"|"+codekg5

lmkg1=request.form("lmkg1")
lmkg2=request.form("lmkg2")
lmkg3=request.form("lmkg3")
lmkg4=request.form("lmkg4")
lmkg5=request.form("lmkg5")
lmkg6=request.form("lmkg6")
lmkg7=request.form("lmkg7")
lmkg8=request.form("lmkg8")
lmkg9=request.form("lmkg9")
lmkg10=request.form("lmkg10")
lmkg11=request.form("lmkg11")
lmkg12=request.form("lmkg12")
lmkg13=request.form("lmkg13")
lmkg14=request.form("lmkg14")
lmkg15=request.form("lmkg15")
If request.form("lmkg1")="" Then lmkg1="0"
If request.form("lmkg2")="" Then lmkg2="0"
If request.form("lmkg3")="" Then lmkg3="0"
If request.form("lmkg4")="" Then lmkg4="0"
If request.form("lmkg5")="" Then lmkg5="0"
If request.form("lmkg6")="" Then lmkg6="0"
If request.form("lmkg7")="" Then lmkg7="0"
If request.form("lmkg8")="" Then lmkg8="0"
If request.form("lmkg9")="" Then lmkg9="0"
If request.form("lmkg10")="" Then lmkg10="0"
If request.form("lmkg11")="" Then lmkg11="0"
If request.form("lmkg12")="" Then lmkg12="0"
If request.form("lmkg13")="" Then lmkg13="0"
If request.form("lmkg14")="" Then lmkg14="0"
If request.form("lmkg15")="" Then lmkg15="0"
lmkg=lmkg1+"|"+lmkg2+"|"+lmkg3+"|"+lmkg4+"|"+lmkg5+"|"+lmkg6+"|"+lmkg7+"|"+lmkg8+"|"+lmkg9+"|"+lmkg10+"|"+lmkg11+"|"+lmkg12+"|"+lmkg13+"|"+lmkg14+"|"+lmkg15

shows1=request.form("shows1")
shows2=request.form("shows2")
shows3=request.form("shows3")
shows4=request.form("shows4")
shows5=request.form("shows5")
shows6=request.form("shows6")
shows7=request.form("shows7")
shows8=request.form("shows8")
shows9=request.form("shows9")
shows10=request.form("shows10")
shows11=request.form("shows11")
shows12=request.form("shows12")
If request.form("shows1")="" Then shows1="0"
If request.form("shows2")="" Then shows2="0"
If request.form("shows3")="" Then shows3="0"
If request.form("shows4")="" Then shows4="0"
If request.form("shows5")="" Then shows5="0"
If request.form("shows6")="" Then shows6="0"
If request.form("shows7")="" Then shows7="0"
If request.form("shows8")="" Then shows8="0"
If request.form("shows9")="" Then shows9="0"
If request.form("shows10")="" Then shows10="0"
If request.form("shows11")="" Then shows11="0"
If request.form("shows12")="" Then shows12="0"
shows=shows1+"|"+shows2+"|"+shows3+"|"+shows4+"|"+shows5+"|"+shows6+"|"+shows7+"|"+shows8+"|"+shows9+"|"+shows10+"|"+shows11+"|"+shows12

Dim certificate,certificate1,certificate2,certificate3,certificate4,certificate5,certificate6,certificate7
certificate1=request.form("certificate1")
certificate2=request.form("certificate2")
certificate3=request.form("certificate3")
certificate4=request.form("certificate4")
certificate5=request.form("certificate5")
certificate6=request.form("certificate6")
certificate7=request.form("certificate7")
If request.form("certificate1")="" Then certificate1="0"
If request.form("certificate2")="" Then certificate2="0"
If request.form("certificate3")="" Then certificate3="0"
If request.form("certificate4")="" Then certificate4="0"
If request.form("certificate5")="" Then certificate5="0"
If request.form("certificate6")="" Then certificate6="0"
If request.form("certificate7")="" Then certificate7="0"
certificate=certificate1+"|"+certificate2+"|"+certificate3+"|"+certificate4+"|"+certificate5+"|"+certificate6+"|"+certificate7

HomeElite1=request.form("HomeElite1")
HomeElite2=request.form("HomeElite2")
HomeElite3=request.form("HomeElite3")
HomeElite4=request.form("HomeElite4")
HomeElite5=request.form("HomeElite5")
HomeElite=HomeElite1+"|"+HomeElite2+"|"+HomeElite3+"|"+HomeElite4+"|"+HomeElite5

Call OpenConn
	set rs=server.createobject("adodb.recordset")
	sql="select * from [DNJY_config] where id="&cstr(id)
	rs.open sql,conn,1,3
    if request("ok")=1 Then
	For I=0 To 12
	S=S&Replace(HtmlEncode(Request.form("S"&i)),"|",",")&"|"
	Next
	S=Left(S,Len(S)-1)
	Application.Lock
	Application(CacheName&"_WebSetting")=Split(S,"|")
	Application.unLock	    
rs("WebSetting")=S
rs("title")=Application(CacheName&"_WebSetting")(0)'request("title")
rs("web")=request("web")
rs("InstallDir")=trim(request("InstallDir"))
rs("logo")=request("logo")
rs("CacheName")=trim(request("CacheName"))
rs("mymail")=Application(CacheName&"_WebSetting")(4)'request("mymail")
rs("kefu")=request("kefu")
rs("kefuid")=request("kefuid")
'rs("qqskin")=request("qqskin")
rs("qqshow")=request("qqshow")
rs("qqimg")=request("qqimg")
rs("kefuskin")=request("kefuskin")
rs("k54kefu")=request("k54kefu")
rs("stopsm")=trim(request("stopsm"))
rs("icp")=request("icp")
rs("synopis")=trim(request("synopis"))
rs("leixing")=request("leixing")
rs("z_hb")=request("z_hb")
rs("z_a")=request("z_a")
rs("z_b")=request("z_b")
rs("z_c")=request("z_c")
rs("z_d")=request("z_d")
rs("jf_1")=request("jf_1")
rs("jf_2")=request("jf_2")
rs("jf_3")=request("jf_3")
rs("jf_4")=request("jf_4")
rs("jf_5")=request("jf_5")
rs("jf_hb")=request("jf_hb")
rs("jf_ck")=request("jf_ck")
rs("adjfs")=request("adjfs")
rs("g_a")=request("g_a")
rs("g_b")=request("g_b")
rs("g_c")=request("g_c")
rs("g_d")=request("g_d")
rs("rmb_mc")=request("rmb_mc")
rs("rmb_hb")=1
rs("vipje")=request("vipje")
rs("huiyuansp")=request("huiyuansp")
rs("huiyuanxx")=request("huiyuanxx")
rs("jiaoyi")=request("jiaoyi")
rs("overdue")=request("overdue")
rs("addxinxi")=request("addxinxi")
rs("xinxish")=request("xinxish")
rs("usernews")=request("usernews")
rs("userlink")=request("userlink")
rs("adjf")=request("adjf")
rs("vipsj")=request("vipsj")
rs("vipno")=request("vipno")
rs("onoff")=request("onoff")
rs("b_y")=request("b_y")
rs("tui_y")=request("tui_y")
rs("xinxis")=request("xinxis")
rs("codekg")=codekg
rs("lmkg")=lmkg
rs("certificate")=certificate
rs("prompt")=request("prompt")
rs("cip")=request("cip")
rs("sms_user")=request("sms_user")
rs("sms_pass")=request("sms_pass")
rs("sms_content")=trim(request("sms_content"))
rs("sms_kg")=request("sms_kg")
rs("sms_cs")=request("sms_cs")
rs("sms_ip")=request("sms_ip")
rs("sms_time")=request("sms_time")
rs("sms_del")=request("sms_del")
rs("shows")=shows
rs("zcfbsj")=request("zcfbsj")
rs("hdlb")=request("hdlb")
rs("contribute")=request("contribute")
rs("slidelx")=request("slidelx")
rs("biaotinum")=request("biaotinum")
rs("memonum")=request("memonum")
rs("addlink")=request("addlink")
rs("linklogo")=request("linklogo")
rs("HomeElite")=HomeElite
rs("hotbz")=request("hotbz")
rs("SY_Type")=request("SY_Type")
rs("SY_interval")=request("SY_interval")
rs("SY_Text")=request("SY_Text")
rs("SY_FontName")=request("SY_FontName")
rs("SY_FontSize")=request("SY_FontSize")
rs("SY_FontColor")=request("SY_FontColor")
rs("SY_Bold")=request("SY_Bold")
rs("SY_FileName")=request("SY_FileName")
rs("SY_Width")=request("SY_Width")
rs("SY_X")=request("SY_X")
rs("SY_Y")=request("SY_Y")
rs("SY_Transparence")=request("SY_Transparence")
rs("SY_BackgroundColor")=request("SY_BackgroundColor")
rs("SY_coordinates")=request("SY_coordinates")
rs("regmail")=request("regmail")
rs("mailqr")=request("mailqr")
rs("usersh")=request("usersh")
rs("zdsh")=request("zdsh")
rs("regyxq")=request("regyxq")
rs("citys")=request("citys")
rs("asphtm")=request("asphtm")
rs("content_length")=request("content_length")
rs("HtmlHome")=request("HtmlHome")
rs("tjdm")=request("tjdm")
rs("Filterhtm")=request("Filterhtm")
rs("gxkf")=request("gxkf")
rs("newspl")=request("newspl")
rs("tgjf")=request("tgjf")
rs("webeditor")=request("webeditor")
rs("plsh")=request("plsh")
rs.update

 dim fso,hf
 set fso=Server.CreateObject(Script_FSO)
 set hf=fso.CreateTextFile(Server.mappath("../Database.asp"),true)
 hf.write "<" & "%" & vbcrlf
 hf.write "'=====���²������뱣����ȷ���ɺ�̨�������ɣ����ü��±��ֹ��޸ģ������׾Ͳ�Ҫ�ҸĶ�====www.ip126.com===" & vbcrlf 
 hf.write vbcrlf
 hf.write "'==========================���ݿ�����====================================" & vbcrlf
 hf.write "Const SystemDatabaseType=" & trim(request("SystemDatabaseType")) & " '���ݿ����� ACCESS���ݿ�Ϊ0 SQL���ݿ�Ϊ1 " & vbcrlf
 hf.write vbcrlf
 hf.write "'==========================Access���ݿ�==================================" & vbcrlf
 hf.write "Const DBFileName=" & chr(34) & trim(request("DBFileName")) & chr(34) & " 'ACCESS���ݿ��ŵ�·��, ����ڸ�Ŀ¼��·��" & vbcrlf
 hf.write vbcrlf
 hf.write "'==========================SQL���ݿ�=====================================" & vbcrlf
 hf.write "Const SqlDatabaseName=" & chr(34) & trim(request("SqlDatabaseName")) & chr(34) & " 'SQL���ݿ���(SqlDatabaseName)" & vbcrlf
 hf.write "Const SqlUsername=" & chr(34) & trim(request("SqlUsername")) & chr(34) & " 'SQL���ݿ��¼�û���(SqlUsername)" & vbcrlf
 hf.write "Const SqlPassword=" & chr(34) & trim(request("SqlPassword")) & chr(34) & "'SQL���ݿ��û���¼����(SqlPassword)" & vbcrlf
 hf.write "Const SqlHostIP=" & chr(34) & trim(request("SqlHostIP")) & chr(34) & "" & vbcrlf
 hf.write "'SQL������ַ(SqlHostIP)����ע�⣺���������ݿ�ͬ����(local)��Զ����IP����""219.48.165.56""�����ݿ��ַΪ��������""���ݿ�����""�����ص���ʱ����SQL Server�����汾���氲װʱ��SQL������ַҪ������д��""�������\ʵ����""����""user\sql2008""" & vbcrlf
 
 hf.write "%" & ">" & vbcrlf
 hf.close
 set hf=nothing
 set fso=Nothing
'================��������ҳ��̬ҳʱɾ�������ɵľ�̬��ҳ�ļ�
If request("HtmlHome")=0 Then
Dim objFSO
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if objFSO.fileExists(Server.MapPath("../index..html")) then
   objFSO.DeleteFile(Server.MapPath("../index..html"))
end if
set objfso=Nothing
End if
'===============================���ɿͷ�����
Dim JsCode,JsFileName
JsCode=Html2Js(request("k54kefu"))
JsCode="document.write('"&JsCode&"')"
JsCode = "<!--" & vbCrLf & JsCode & vbCrLf &"//-->" 
JsFileName = Server.MapPath("..\UploadFiles\js\54kefu.js")
Call WriteToFile(JsFileName,JsCode,"utf-8")'�ú�����ʽ

url="admin_parameters.asp?id=1"
response.Write "<script language=javascript>alert('�����ɹ������Զ����������ļ�config.asp');location.href='admin_config.asp?page="&url&"';</script>"
response.end
else
%>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<script language="javascript">
function changedbtype(dbtype){
  var accesstr=document.getElementById("accesstr");
  var sqltr=document.getElementById("sqltr");
  if(dbtype==0){
  	accesstr.style.display = '';
	sqltr.style.display = 'none';
  }else{
  	accesstr.style.display = 'none';
	sqltr.style.display = '';
  }
}
</script>
<SCRIPT language="javascript" src="admin.js"></SCRIPT>
<IFRAME width="260" height="165" id="colourPalette" src="nc_selcolor.htm" style="visibility:hidden; position: absolute; left: 0px; top: 0px;border:1px gray solid" frameborder="0" scrolling="no" ></IFRAME>
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
}
.style1 {
	color: #0000FF;
	font-weight: bold;
}
.style2 {color: #FF0000}
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
<script language="javascript">
<!--
//˵���ֽ������ƣ����������=========
String.prototype.Length=function(){
temp=this.replace(/([^\x00-\xff])/g,"$1$1");
return temp.length
}
//��ʾ�����ַ���
function kn()
  {
   var ln=myform.prompt.value.Length();
     num.innerText='������������:'+(200-ln)+'���ַ�'+(<%=200/2%>-ln/2)+'����';
      //if (ln>=8) form.memo.readOnly=true;  //��������������������޷�������޸�
}

function myform_onsubmit() {
if (document.myform.S0.value.length>50)
	{
	  alert("��վ���Ʋ��ܳ���50���ַ�")
	  document.myform.S0.focus()
	  return false
	 }
if (document.myform.web.value.length>50)
	{
	  alert("��վ�������ܳ���50���ַ�")
	  document.myform.web.focus()
	  return false
	 }

function contain(str,charset)//�ַ����������Ժ���
{
//�����������ִ��в��ܰ���ָ���ַ��е��κ�һ��
//var i;
//for(i=0;i<charset.length;i++)
//if(str.indexOf(charset.charAt(i))>=0)
if(str.indexOf(charset)>=0)//����Ϊ�ִ��в��ܰ���ָ���ַ�������
return true;
return false;
} 
if(contain(document.myform.web.value,"http://"))
{ 
alert("��ַǰ��Ҫ��http://");
document.myform.web.focus();
return false;
}
if(contain(document.myform.web.value,"/"))
{ 
alert("��ַ��Ҫ��/");
document.myform.web.focus();
return false;
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
    if ((document.myform.S4.value!="")&&(!checkemail(myform.S4.value)))
	{
    alert("������Email��ַ����ȷ������������!");
    document.myform.S4.focus();
    return false;
    }

if (document.myform.prompt.value.length>200)
	{
	  alert("�������������������ݲ��ܳ���200���ַ�")
	  document.myform.prompt.focus()
	  return false
	 }		  

}

// --></script>
</head>
<BODY>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse">
  <tr>
    <td align="center">
      <form name="myform" method="POST" id="myform" language="javascript" onSubmit="return myform_onsubmit()" action="admin_parameters.asp?ok=1&id=<%=request("id")%>">
        <table border="1" cellspacing="0" cellpadding="0" width="100%" style="border-collapse: collapse" bordercolor="#ADAED6" height="799">
          <tr> 
            <td width="688" height="28" bgcolor="#BDBEDE">
            <p class="style1"><font size="2">&nbsp;���������趨</font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" align="center" style="border-left-style: none; border-left-width: medium; border-right-style: none; border-right-width: medium" width="700" height="735"> 
              <table width="100%" border="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
              
                <tr> 
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium; border-top-style:solid; border-top-width:1" height="27" bgcolor="#FFFFFF" align="right">�趨��վ���ƣ�</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium; border-top-style:solid; border-top-width:1" height="27" bgcolor="#FFFFFF"> 
                    <p style="margin-left: 20"> 
                    <input type="text" name="S0" value="<%=Application(CacheName&"_WebSetting")(0)%>" size="40"> 
                    </td>
                </tr>
                <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium; border-top-style:none; border-top-width:medium" height="27" bgcolor="#FFFFFF" align="right">��վ������</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium; border-top-style:none; border-top-width:medium" height="27" bgcolor="#FFFFFF"> 
                    <p style="margin-left: 20"> 
                    <input type="text" name="web" value="<%=rs("web")%>" size="40">(ǰ�治����http://�������治����/�������ڸ�Ŀ¼���¼�Ŀ¼��װ��ϵͳ�����治Ҫ����Ŀ¼��)
                    </td>
                </tr>
                <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium; border-top-style:none; border-top-width:medium" height="27" bgcolor="#FFFFFF" align="right"><font color="#FF0000">ϵͳ��װĿ¼(��Ҫ)��</font></td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium; border-top-style:none; border-top-width:medium" height="27" bgcolor="#FFFFFF"> 
                    <p style="margin-left: 20"> 
                    <input type="text" name="InstallDir" value="<%=strInstallDir%>" size="40"> &nbsp;<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>(����ڸ�Ŀ¼��λ�ã�ǰ�󶼲�����/��,���ڸ�Ŀ¼�����ա�Ŀ¼���ĺ�Ҫ������������ģ�塢htm�ļ��͹��ͼƬ·���������)</span></a></td>
                </tr>
				 <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium; border-top-style:none; border-top-width:medium" height="27" bgcolor="#FFFFFF" align="right">���ݿ����ͣ�</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium; border-top-style:none; border-top-width:medium"  bgcolor="#FFFFFF"><p style="margin-left: 20"><input name="SystemDatabaseType" onClick="javascript:changedbtype(0);" type="radio" value="0" <%if SystemDatabaseType=0 then%> checked<%end if%>>ACCESS
<input type="radio" onClick="javascript:changedbtype('1');" name="SystemDatabaseType" value="1" <%if SystemDatabaseType="1" then%> checked<%end if%>>SQL��������ҵ�棩����SQL�洢���̻��и��õ����ܱ���<br><font color="#FF0000">���ݰ汾��ȷѡ�����������裬�����޷��������ݿ⣡</font></td>
                </tr>

<tr id="accesstr" name="accesstr"> 
<td width="30%" style="border-bottom-style: none; border-bottom-width: medium; border-top-style:none; border-top-width:medium" height="27" bgcolor="#FFFFFF" align="right">ACCSS���ݿ��ļ�·����</td>
<td><p style="margin-left: 20"><input name="DBFileName" type="text" id="DBFileName" value="<%=DBFileName%>" size="40" maxlength="50"> <br>��չ����Ϊasp�ɷ����ء�������Ҫ�ֹ��޸���Ӧ��Ŀ¼�����ݿ��ļ���</td>
</tr>
<tr id="sqltr" name="sqltr"> 
<td width="30%" style="border-bottom-style: none; border-bottom-width: medium; border-top-style:none; border-top-width:medium" height="27" bgcolor="#FFFFFF" align="right">MS SQL���ݿ����ã�</td>
<td colspan="2"><p style="margin-left: 20"><table  height="80" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td><input name="SqlDatabaseName" type="text" id="SqlDatabaseName" value="<%=SqlDatabaseName%>" size="25" maxlength="50"> SQL���ݿ���</td>
  </tr>
  <tr>
    <td><input name="SqlUsername" type="password" id="SqlUsername" value="<%=SqlUsername%>" size="25" maxlength="50"> ��¼���ݿ��û���</td>
  </tr>
  <tr>
    <td><input name="SqlPassword" type="password" id="SqlPassword" value="<%=SqlPassword%>" size="25" maxlength="50"> ��¼���ݿ�����</td>
  </tr>
  <tr>
    <td><input name="SqlHostIP" type="text" id="SqlLocalName" value="<%=SqlHostIP%>" size="25" maxlength="50"> ��������Դ[վ��ͬ����(local)��վ�������IP������Դ����]</td>
  </tr>
</table></td>
</tr>
<script language="javascript">changedbtype(<%=SystemDatabaseType%>);</script>				
                <tr> 
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">��վ��������LOGO��</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="logo" value="<%=rs("logo")%>" size="40"> 
                    ����С88*31��gif��ʽ��</td>
                </tr>
                <tr> 
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">��վ��Ӫ�����ţ�</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="icp" value="<%=rs("icp")%>" size="40"> 
                    </td>
                </tr>
                 <tr> 
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">��˾���ƣ�</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="S12" value="<%=Application(CacheName&"_WebSetting")(12)%>" size="40"> 
                    </td>
                </tr>
                 <tr> 
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">�칫��ַ��</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="S8" value="<%=Application(CacheName&"_WebSetting")(8)%>" size="40"> 
                    </td>
                </tr>
                 <tr> 
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">�������룺</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="S9" value="<%=Application(CacheName&"_WebSetting")(9)%>" size="30" <%=onKeyUp%>> 
                    </td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">�칫�绰��</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="S5" value="<%=Application(CacheName&"_WebSetting")(5)%>" size="30">
                  </td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">��ϵ�ֻ���</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="S6" value="<%=Application(CacheName&"_WebSetting")(6)%>" size="30" <%=onKeyUp%>>
                  </td>
                </tr>
                 <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">������룺</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="S7" value="<%=Application(CacheName&"_WebSetting")(7)%>" size="30">
                  </td>
                </tr>
                  <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">�������䣺</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="S4" value="<%=Application(CacheName&"_WebSetting")(4)%>" size="30"> (�����ʼ�ϵͳ�����еķ�����������ͬ)
                  </td>
                </tr>               
 

                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">��վֹͣ����˵����</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><textarea name="stopsm" cols="80" rows="3" class="input2"><%=rs("stopsm")%></textarea></td>   
                </tr>
           
                    
                <tr> 
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF" align="right">�����������ã�</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF">
                    <p style="margin-left: 20"><input type="text" name="leixing" value="<%=rs("leixing")%>" size="45"> <br><font color="#FF0000">
����֮���ð�ǵġ�|���ָ���ǰ�󲻴���|������</font>�����һ|����|�����                                      
                    </td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">��վ��������Ϣ��</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><textarea name="S1" cols="80" rows="3" class="input2"><%=Application(CacheName&"_WebSetting")(1)%></textarea></td>
                </tr>
                  <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">��վ�ؼ���(keywords)��<br>���������������Ĺؼ����ݣ��ؼ����á�,���ŷָ�</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><textarea name="S3" cols="80" rows="5" class="input2"><%=Application(CacheName&"_WebSetting")(3)%></textarea></td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">��վ˵��(description)��</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><textarea name="S2" cols="80" rows="5" class="input2"><%=Application(CacheName&"_WebSetting")(2)%></textarea></td>
                </tr>

                 <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">��վ��飺</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><textarea name="synopis" cols="80" rows="5" class="input2"><%=rs("synopis")%></textarea></td>
                </tr>
                 <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">ͳ�ƴ��룺</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><textarea name="tjdm" cols="80" rows="5" class="input2"><%=rs("tjdm")%></textarea> ���й�����ͳ�Ʒ����̴�����һ������ͳ�ƴ������ڴ˲�����β��ģ��Ϳ�ͳ����</td>
                </tr>

                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">�����������룺
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"> <%if rs("addlink")=0 then%>               
                <input type="radio" name="addlink" value="0" id="addlink" checked>
                 ����&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="addlink" value="1" id="addlink">
                ��ֹ
                <%else%>                         
                <input type="radio" name="addlink" value="0" id="addlink">
                 ����&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="addlink" value="1" id="addlink" checked>
                ��ֹ<%end if%> 
                  </td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">���������ϴ�LOGO��
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"> <%if rs("linklogo")=1 then%>               
                <input type="radio" name="linklogo" value="1" id="linklogo" checked>
                 ����&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="linklogo" value="0" id="linklogo">
                ��ֹ
                <%else%>                         
                <input type="radio" name="linklogo" value="1" id="linklogo">
                 ����&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="linklogo" value="0" id="linklogo" checked>
                ��ֹ<%end if%> 
                  </td>
                </tr>
<%
lmkg=rs("lmkg")
lmkg=split(lmkg,"|")
lmkg1=lmkg(0)
lmkg2=lmkg(1)
lmkg3=lmkg(2)
lmkg4=lmkg(3)
lmkg5=lmkg(4)
lmkg6=lmkg(5)
lmkg7=lmkg(6)
lmkg8=lmkg(7)
lmkg9=lmkg(8)
lmkg10=lmkg(9)
lmkg11=lmkg(10)
lmkg12=lmkg(11)
lmkg13=lmkg(12)
lmkg14=lmkg(13)
lmkg15=lmkg(14)
%>
                  <tr bgcolor="#BDBEDE">
                  <td height="28" colspan="2" style="border-top-style: none; border-top-width: medium; border-bottom-style: solid; border-bottom-width: 1">
                    <span class="style1"><font size="2">&nbsp;��ҳ��鿪������</font></span></td>
                </tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">��ҳ��鿪�أ�</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20"><INPUT TYPE=checkbox NAME="lmkg1" VALUE="1" <%if lmkg1="1" then response.write("checked")%>>������վ&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="lmkg2" VALUE="1" <%if lmkg2="1" then response.write("checked")%>>��վֻ��&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="lmkg3" VALUE="1" <%if lmkg3="1" then response.write("checked")%>>���е���&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="lmkg4" VALUE="1" <%if lmkg4="1" then response.write("checked")%>>��Ŀ�б�
					&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="lmkg5" VALUE="1" <%if lmkg5="1" then response.write("checked")%>>������Ϣ
					&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="lmkg6" VALUE="1" <%if lmkg6="1" then response.write("checked")%>>��ҵ��Ƭ<br>
					<INPUT TYPE=checkbox NAME="lmkg7" VALUE="1" <%if lmkg7="1" then response.write("checked")%>>�Ƽ�����
					&nbsp;<INPUT TYPE=checkbox NAME="lmkg8" VALUE="1" <%if lmkg8="1" then response.write("checked")%>>����114					
					&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="lmkg9" VALUE="1" <%if lmkg9="1" then response.write("checked")%>>���ƹ��
					&nbsp;<INPUT TYPE=checkbox NAME="lmkg10" VALUE="1" <%if lmkg10="1" then response.write("checked")%>>��������&nbsp;&nbsp; <INPUT TYPE=checkbox NAME="lmkg11" VALUE="1" <%if lmkg11="1" then response.write("checked")%>>ͼƬ��Ϣ
					&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="lmkg12" VALUE="1" <%if lmkg12="1" then response.write("checked")%>>�������<br>
					<INPUT TYPE=checkbox NAME="lmkg13" VALUE="1" <%if lmkg13="1" then response.write("checked")%>>��Դ����&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="lmkg14" VALUE="1" <%if lmkg14="1" then response.write("checked")%>>�������&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="lmkg15" VALUE="1" <%if lmkg15="1" then response.write("checked")%>>������־
<br>�ɸ�����Ҫ������ҳ���</td>
                </tr>
                  <tr bgcolor="#BDBEDE">
                  <td height="28" colspan="2" style="border-top-style: none; border-top-width: medium; border-bottom-style: solid; border-bottom-width: 1">
                    <span class="style1"><font size="2">&nbsp;������ʾ����</font></span></td>
                </tr>
<%
shows=rs("shows")
shows=split(shows,"|")
shows1=shows(0)
shows2=shows(1)
shows3=shows(2)
shows4=shows(3)
shows5=shows(4)
shows6=shows(5)
shows7=shows(6)
shows8=shows(7)
shows9=shows(8)
shows10=shows(9)
shows11=shows(10)
shows12=shows(11)
%>
				 <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">�Զ����뱾�أ�
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20">
		 <%Dim rscip
		  set rscip=server.CreateObject("adodb.recordset")
		  rscip.Open "select * from DNJY_city where id>0 and twoid=0 order by id",conn,1,1		  
		  if rscip.EOF and rscip.BOF then%>	
          <input type="hidden" name="cip" value="0" id="cip">
			����������ĿǰΪ�գ�������Ч��
         <%else%>
				<%if rs("cip")=True then%>               
                <input type="radio" name="cip" value="1" id="cip" checked>
                 ����&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="cip" value="0" id="cip">
                �ر�
                <%else%>                         
                <input type="radio" name="cip" value="1" id="cip">
                 ����&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="cip" value="0" id="cip" checked>
                �ر�<%end if%> 
		  <%rscip.close
          set rscip=Nothing
          End If%>
                  <span class="style2"><br>ѡ������ʱ�Զ����ݷÿ�IP��ַ�����������ط�վ��IP��ַ���ϸ��ӿ����в�׼ȷ�����������������̫�࣬��ȫ��ʡ������������վ��ȡ���ݺͱȶ�Ҫһ��ʱ�䣬�Ե���վ�ٶ��г������󣬽������ã�������ݵ��������������ضԱ��ٶ�ȷ���Ƿ����á���</span></td>
                </tr>
				<tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">����վ��ʾ��</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20"><INPUT TYPE=checkbox NAME="shows1" VALUE="1" <%if shows1="1" then response.write("checked")%>>������Ϣ&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="shows2" VALUE="1" <%if shows2="1" then response.write("checked")%>>������Ѷ&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="shows3" VALUE="1" <%if shows3="1" then response.write("checked")%>>�̼ҵ���&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="shows4" VALUE="1" <%if shows4="1" then response.write("checked")%>>��ҵ��Ƭ
					&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="shows5" VALUE="1" <%if shows5="1" then response.write("checked")%>>�û�����&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="shows6" VALUE="1" <%if shows6="1" then response.write("checked")%>>����114 <br>
					<INPUT TYPE=checkbox NAME="shows7" VALUE="1" <%if shows7="1" then response.write("checked")%>>������Ϣ
					&nbsp;<INPUT TYPE=checkbox NAME="shows8" VALUE="1" <%if shows8="1" then response.write("checked")%>>��վ���&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="shows9" VALUE="1" <%if shows9="1" then response.write("checked")%>>��������&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="shows10" VALUE="1" <%if shows10="1" then response.write("checked")%>>��վLOGO&nbsp;&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="shows11" VALUE="1" <%if shows11="1" then response.write("checked")%>>�������&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="shows12" VALUE="1" <%if shows12="1" then response.write("checked")%>>������־
					&nbsp;
<br><span class="style2">�����ѡ����ð�鰴��վ��ȫ������������ʾ</span></td>
                </tr> 				
                <tr bgcolor="#BDBEDE">
                  <td height="28" colspan="2" style="border-top-style: none; border-top-width: medium; border-bottom-style: solid; border-bottom-width: 1">
                    <span class="style1"><font size="2">&nbsp;�ϴ�ͼƬˮӡ����</font></span></td>
                </tr>
               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right"></td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20">������������������������Ϊ��Ϣ��������Ѷ�ϴ�ͼƬˮӡ���ã���������������������<br>
				  ��AspJpeg���֧�֣�<%if IsObjInstalled("Persits.Jpeg") then%>��<%else%>��<%end if%>&nbsp;<span class="fontcolor">����������֧��AspJpeg���������ϵͳ���Զ��رձ����ܣ�</span></td>
                </tr>
               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">ˮӡ���ͣ�</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="radio" name="SY_Type" value="0" <%if rs("SY_Type")=0 then%>checked<%end if%>>����ˮӡ	<input type="radio" name="SY_Type" value="1" <%if rs("SY_Type")=1 then%>checked<%end if%>>����ˮӡ	<input type="radio" name="SY_Type" value="2" <%if rs("SY_Type")=2 then%>checked<%end if%>>ͼƬˮӡ	</td>
                </tr>
               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">����ʱ�䣺</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="SY_interval" value="<%=rs("SY_interval")%>" size="10" <%=onKeyUp%>>
                  �����ڱ༭��ͼƬ���ظ���ˮӡ</td>
                </tr>
               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">ˮӡλ�����꣺</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20">X��<input name='SY_X' type='text' value='<%=rs("SY_X")%>' size='10' maxlength='10'  ONKEYPRESS="javascript:event.returnValue=IsDigit()" <%=onKeyUp%>> ����<br>Y��<input name='SY_Y' type='text' value='<%=rs("SY_Y")%>' size='10' maxlength='10'  ONKEYPRESS="javascript:event.returnValue=IsDigit()" <%=onKeyUp%>> ����</td>
                </tr>
               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">����ˮӡ��</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="SY_Text" value="<%=rs("SY_Text")%>" size="30">
                  �����������˳���15���ַ�����֧���κ�WEB������</td>
                </tr>
               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">�������壺</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20">
			<SELECT name="SY_FontName" >
            <option value="����"  <% if rs("SY_FontName")="����" Then Response.write("Selected") %>>����</option>
            <option value="����_GB2312" <% if rs("SY_FontName")="����_GB2312" Then Response.write("Selected") %>>����</option>
            <option value="����_GB2312" <% if rs("SY_FontName")="����_GB2312" Then Response.write("Selected") %>>������</option>
            <option value="����" <% if rs("SY_FontName")="����" Then Response.write("Selected") %>>����</option>
            <option value="����" <% if rs("SY_FontName")="����" Then Response.write("Selected") %>>����</option>
            <option value="��Բ" <% if rs("SY_FontName")="��Բ" Then Response.write("Selected") %>>��Բ</option>
            <option value="Andale Mono" <% if rs("SY_FontName")="Andale Mono" Then Response.write("Selected") %>>Andale Mono</OPTION> 
            <option value="Arial"  <% if rs("SY_FontName")="Arial" Then Response.write("Selected") %>>Arial</OPTION> 
            <option value="Arial Black"  <% if rs("SY_FontName")="Arial Black" Then Response.write("Selected") %>>Arial Black</OPTION> 
            <option value="Book Antiqua"  <% if rs("SY_FontName")="Book Antiqua" Then Response.write("Selected") %>>Book Antiqua</OPTION>

        </SELECT></td>
                </tr>
               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">���ִ�С��</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="SY_FontSize" value="<%=rs("SY_FontSize")%>" size="10" <%=onKeyUp%>>
                  ����</td>
                </tr>
               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">������ɫ��</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><INPUT name="SY_FontColor" type="text" style="background:<%=rs("SY_FontColor")%>" onClick="Getcolor(this)" value="<%=rs("SY_FontColor")%>" size="7" maxlength="7"></td>
                </tr>
               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">�Ƿ���壺</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><SELECT name='SY_Bold' >
				  <OPTION value='1' <% if rs("SY_Bold")=1 Then Response.write("Selected") %>>��</OPTION>
                  <OPTION value='0' <% if rs("SY_Bold")=0 Then Response.write("Selected") %>>��</OPTION>
            
          </SELECT></td>
                </tr>


               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">ͼƬˮӡ�ļ�����</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input name='SY_FileName' type='text' value='<%=rs("SY_FileName")%>' size='35' maxlength='40'> ������дͼƬ�ļ������·�����ԡ�\����ͷ</td>
                </tr>
               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">ͼƬˮӡ��ȣ�</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="SY_Width" value="<%=rs("SY_Width")%>" size="10" <%=onKeyUp%>>
                  ���أ�����ݿ���Զ������߶�,����150*60͸����</td>
                </tr>
 
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">ͼƬˮӡ͸���ȣ�</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="SY_Transparence" value="<%=rs("SY_Transparence")%>" size="10" <%=onKeyUp%>>
                  0��1��Χ��1��ʾ��͸��</td>
                </tr>
				  <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">ͼƬˮӡ��ɫ��</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><INPUT name="SY_BackgroundColor" type="text" style="background:<%=rs("SY_BackgroundColor")%>" onClick="Getcolor(this)" value="<%=rs("SY_BackgroundColor")%>" size="7" maxlength="7">
                  ����ȥ��ˮӡͼƬ�ĵ�ɫ�����ڴ������ɫ��RGBֵ</td>
                </tr>
               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">ͼƬˮӡ������㣺</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><select name="SY_coordinates" size="1">
	  <option value="0" <%if rs("SY_coordinates")=0 then%>selected="selected"<%end if%>>����</option>
	  <option value="1" <%if rs("SY_coordinates")=1 then%>selected="selected"<%end if%>>����</option>
	  <option value="2" <%if rs("SY_coordinates")=2 then%>selected="selected"<%end if%>>����</option>
	  <option value="3" <%if rs("SY_coordinates")=3 then%>selected="selected"<%end if%>>����</option>
	  <option value="4" <%if rs("SY_coordinates")=4 then%>selected="selected"<%end if%>>����</option>
	  </select></td>
                </tr>
                <tr bgcolor="#BDBEDE">
                  <td height="28" colspan="2" style="border-top-style: none; border-top-width: medium; border-bottom-style: solid; border-bottom-width: 1">
                    <span class="style1"><font size="2">&nbsp;���߿ͷ��趨</font></span></td>
                </tr>

               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">�ͷ����ͣ�</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><select name="kefu" style="BACKGROUND-COLOR: #FFCCFF; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;">
                            <option value="0" <% if rs("kefu")=0 Then Response.write("Selected") %>>�رտͷ�</option>                            
                            <option value="1" <% if rs("kefu")=1 Then Response.write("Selected") %>>��ѶQQ</option>
							<option value="2" <% if rs("kefu")=2 Then Response.write("Selected") %>>53���</option>
							<option value="3" <% if rs("kefu")=3 Then Response.write("Selected") %>>54�ͷ�</option>
                            <option value="4" <% if rs("kefu")=4 Then Response.write("Selected") %>>53���+��ѶQQ</option>
                          </select>
                  </td>
                </tr>

               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">QQ���룺</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="S10" value="<%=Application(CacheName&"_WebSetting")(10)%>" size="30">
                  ���QQ�������Ķ��Ÿ�����QQ����������ǳ�һһ��Ӧ</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">QQ��Ӧ�ǳƣ�</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="S11" value="<%=Application(CacheName&"_WebSetting")(11)%>" size="30">
                  ����ǳ������Ķ��Ÿ������ǳ��������QQ��һһ��Ӧ</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">QQ��ʾ��Ŀ��</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="qqshow" value="<%=rs("qqshow")%>" size="30" <%=onKeyUp%>>
                  0QQ�ţ�1�ǳƣ�3ͼ�꣬4QQ��+�ǳƣ�5�ǳ�+ͼ�꣬6QQ��+�ǳ�+ͼ��</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">QQͼ����</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="qqimg" value="<%=rs("qqimg")%>" size="30" <%=onKeyUp%>>
                  1��13�֣�<a target=blank href='http://www.tool.la/QQCode/'><font color="#0000FF">��ʽ�鿴</font></a></td>
                </tr>
                <!--tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">QQ��ʾλ�ã�</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20">
				  <%if rs("qqskin")=1 then%>               
                <input type="radio" name="qqskin" value="1" id="qqskin" checked>
                 ���&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="qqskin" value="0" id="qqskin">
                �Ҳ�
                <%else%>                         
                <input type="radio" name="qqskin" value="1" id="qqskin">
                 ���&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="qqskin" value="0" id="qqskin" checked>
                �Ҳ�<%end if%>&nbsp; <img align=top src=../images/memo.gif alt="�ͷ�QQ����Ļ��λ��">
                 </td>
                </tr-->

                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">QQ��ʾ����ʽ��</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="radio" name="kefuskin" value="1" <%if rs("kefuskin")=1 then%>checked<%end if%>>��ʽһ	<input type="radio" name="kefuskin" value="2" <%if rs("kefuskin")=2 then%>checked<%end if%>>��ʽ��	<input type="radio" name="kefuskin" value="3" <%if rs("kefuskin")=3 then%>checked<%end if%>>��ʽ��	<input type="radio" name="kefuskin" value="4" <%if rs("kefuskin")=4 then%>checked<%end if%>>��ʽ��	<input type="radio" name="kefuskin" value="5" <%if rs("kefuskin")=5 then%>checked<%end if%>>��ʽ��<a class="info" href=../images/qq/qq.gif target=_blank><img src="../images/memo.gif"  BORDER="0" /><span>���Ԥ��Ч��</span></a></td>
                </tr>              
                 <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">53����ʺţ�</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="kefuid" value="<%=rs("kefuid")%>" size="30">
                  <br>��<a target=blank href='http://www.53kf.com'><font color="#0000FF">53�������</font></a>����һ���ʺŲ����ذ�װ�ͷ��˼��ɣ��谲װ�ͻ��ˣ������Ѻã�</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">54�ͷ����룺</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><textarea NAME="k54kefu" ROWS=3 style="overflow:auto;width=70%"><%=rs("k54kefu")%></textarea>
                  <br>��<a target=blank href='http://wwww.54kefu.net/'><font color="#0000FF">54�ͷ�����</font></a>����һ���ʺŲ���ȡ�������ϼ��ɣ��ɼ��ɶ��ֿͷ����ߣ����ɷ��ӣ�</td>
                 </tr>


                <tr bgcolor="#BDBEDE">
                  <td height="28" colspan="2" style="border-top-style: none; border-top-width: medium; border-bottom-style: solid; border-bottom-width: 1">
                    <span class="style1"><font size="2">&nbsp;����������������</font></span></td>
                </tr>
                <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">����������ƣ�</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="rmb_mc" size="20" value="<%=rs("rmb_mc")%>">
                    ������վ�ص��������Ҹ������ƣ�������ҡ�ͭ�桢���</td>
                </tr> 
                <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">
                  <font color="#0000FF">1Ԫ�����</font>����<font color="#0000FF">����</font>��</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="rmb_hb" size="20" value="<%=rs("rmb_hb")%>" DISABLED <%=onKeyUp%>>
                    ��<%=rs("rmb_mc")%>��Ϊ��������ҵĹ���ת�������ֵ��Ҫ�޸��ˣ�</td>
                </tr> 

                <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">����<font color="#990000">������ɫ����</font>�裺</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="g_a" size="20" value="<%=rs("g_a")%>" <%=onKeyUp%>>
                    ��<%=rs("rmb_mc")%></td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">����<font color="#990000">��Ϣ�ö�����</font>�裺</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="g_b" size="20" value="<%=rs("g_b")%>" <%=onKeyUp%>>
                    ��<%=rs("rmb_mc")%></td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">����<font color="#990000">������ͼ����</font>�裺</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="g_c" size="20" value="<%=rs("g_c")%>" <%=onKeyUp%>>
                    ��<%=rs("rmb_mc")%></td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">����<font color="#990000">ͨ����˵���</font>�裺</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="g_d" size="20" value="<%=rs("g_d")%>" <%=onKeyUp%>>
                    ��<%=rs("rmb_mc")%></td>
                </tr>
                <tr bgcolor="#BDBEDE">
                  <td height="28" colspan="2" style="border-top-style: none; border-top-width: medium; border-bottom-style: solid; border-bottom-width: 1">
                    <span class="style1"><font size="2">&nbsp;���ֽ�����ת������</font></span></td>
                </tr>
                <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">ע�����ͣ�</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="jf_1" size="20" value="<%=rs("jf_1")%>" <%=onKeyUp%>>
                    �����</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">������Ϣÿ�����ͣ�</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="jf_2" size="20" value="<%=rs("jf_2")%>" <%=onKeyUp%>>
                    ����֣��Լ�ɾ����ͬ���֣�����Աɾ���ӱ��۷֣�</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">������ѶͶ��ÿ�����ͣ�</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="tgjf" size="20" value="<%=rs("tgjf")%>" <%=onKeyUp%>>
                    �����</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">��½һ�λ�ȡ��</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="jf_3" size="20" value="<%=rs("jf_3")%>" <%=onKeyUp%>>
                    ����� (ÿ��һ��)</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">����<%=rs("rmb_mc")%>�����ӣ�</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="jf_4" size="20" value="<%=rs("jf_4")%>" <%=onKeyUp%>>
                    �����</td>
                </tr>

                <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">����ת��<%=rs("rmb_mc")%>��</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="jf_hb" size="20" value="<%=rs("jf_hb")%>" <%=onKeyUp%>>
                     ����ֿ�ת��1��<%=rs("rmb_mc")%></td>
                </tr> 
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">1��<%=rs("rmb_mc")%>ת����</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="adjfs" size="20" value="<%=rs("adjfs")%>" <%=onKeyUp%>>
                    �����</td>
                </tr>
                <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">����ע���Ƽ��ˣ�</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="jf_5" size="20" value="<%=rs("jf_5")%>" <%=onKeyUp%>>
                    �����</td>
                </tr>
                <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">�鿴Ȩ����Ϣ��ϵ�ֻ��ۣ�</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="jf_ck" size="20" value="<%=rs("jf_ck")%>" <%=onKeyUp%>>
                    �����</td>
                </tr>                <tr bgcolor="#BDBEDE">
                  <td height="28" colspan="2" bgcolor="#BDBEDE" style="border-top-style: none; border-top-width: medium; border-bottom-style: solid; border-bottom-width: 1">
                    <span class="style1"><font size="2">&nbsp;��Աע��ѡ��</font></span></td>
                </tr>
                <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">���͵�<font color="#0000FF"><%=rs("rmb_mc")%></font><font color="#800000">��</font></td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="z_hb" size="20" value="<%=rs("z_hb")%>" <%=onKeyUp%>>
                    ��</td>
                </tr>
                <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">���͵�<font color="#800000">�����ɫ���ߣ�</font></td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="z_a" size="20" value="<%=rs("z_a")%>" <%=onKeyUp%>>
                    ��</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">���͵�<font color="#800000">��Ϣ�ö����ߣ�</font></td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="z_b" size="20" value="<%=rs("z_b")%>" <%=onKeyUp%>>
                    ��</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">���͵�<font color="#800000">������ͼ���ߣ�</font></td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="z_c" size="20" value="<%=rs("z_c")%>" <%=onKeyUp%>>
                    ��</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">���͵�<font color="#800000">ͨ����˵��ߣ�</font></td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="z_d" size="20" value="<%=rs("z_d")%>" <%=onKeyUp%>>
                    �� �����ظ���˵��ߣ�</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">ע��ͬʱ�����ʼ�֪ͨ��</font></td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="regmail" size="20" value="<%=rs("regmail")%>" <%=onKeyUp%>> ��0=�����ͣ�1=���ͣ�</td>
                </tr>
				                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">ע���Ա�Ƿ�Ҫ�ʼ�ȷ�ϣ�</font></td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="mailqr" size="20" value="<%=rs("mailqr")%>" <%=onKeyUp%>> ��0=����Ҫ��1=��Ҫ��</td>
                </tr>
				  <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">ע���Ա�Ƿ�Ҫ��ˣ�</font></td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="usersh" size="20" value="<%=rs("usersh")%>" <%=onKeyUp%>> ��0=����ˣ�1=��ˣ�</td>
                </tr>
				  <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">�ʼ�ȷ�ϵ�ͬʱ�Զ���ˣ�</font></td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="zdsh" size="20" value="<%=rs("zdsh")%>" <%=onKeyUp%>> ��0=���Զ���ˣ�1=�Զ���ˣ�</td>
                </tr>
				  <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">�ʼ�ȷ����Чʱ�䣺</font></td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="regyxq" size="20" value="<%=rs("regyxq")%>" <%=onKeyUp%>>�� �������������Զ�ɾ����ʱע�����ϣ�</td>
                </tr>
				 <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">�ֻ�������֤���أ�
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"> <%if rs("sms_kg")=True then%>               
                <input type="radio" name="sms_kg" value="1" id="sms_kg" checked>
                 ����&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="sms_kg" value="0" id="sms_kg">
                �ر�
                <%else%>                         
                <input type="radio" name="sms_kg" value="1" id="sms_kg">
                 ����&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="sms_kg" value="0" id="sms_kg" checked>
                �ر�<%end if%> �����û���ʺ�Ҫ�رգ������Ա�޷�ע�ᣩ</td>
                </tr>
				  <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">�ֻ�������֤�ʺţ�</font></td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="sms_user" size="20" value="<%=rs("sms_user")%>"> ����<a target=blank href='http://www.dxton.com/'><font color="#0000FF">����ͨ</font></a>����һ���ʺŲ���д�йز�����<a target=blank href='http://www.dxton.com/webservice/sms.asmx/GetNum?account=<%=rs("sms_user")%>&password=<%=rs("sms_pass")%>'><font color="#0000FF">��ѯ�ʻ����</font></a>��</td>
                </tr>
				  <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">�ֻ�������֤���룺</font></td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="password" name="sms_pass" size="20" value="<%=rs("sms_pass")%>"> ��������Ҫ�����ͨƽ̨�Ľӿ�����һ�£�</td>
                </tr> 
                 <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">�ֻ�������֤���ݣ�</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20">������֤����:****��<textarea name="sms_content" cols="55" rows="2" class="input2"><%=rs("sms_content")%></textarea> <br>��<span class="style2">��Ҫ�����������ֻ���֤�ʺŲ���VIP3���ϵģ������ݾ��Բ����������޸��κ��ַ��������޷�ͨ������ͨƽ̨��˶��޷����Ͷ��ţ�</span>��</td>
                </tr>
				 <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">�ֻ����ܷ��ظ�ע�᣺
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"> <%if rs("sms_cs")=True then%>               
                <input type="radio" name="sms_cs" value="1" id="sms_cs" checked>
                 ����&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="sms_cs" value="0" id="sms_cs">
                ����
                <%else%>                         
                <input type="radio" name="sms_cs" value="1" id="sms_cs">
                 ����&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="sms_cs" value="0" id="sms_cs" checked>
                ����<%end if%> �����鲻׼�ظ�ע�ᣩ</td>
                </tr>
                 <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">ͬһIPע����֤������</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="sms_ip" size="5" value="<%=rs("sms_ip")%>" <%=onKeyUp%>> (0Ϊ������)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ȡ��ͬIP��֤��������ʱ�� <input type="text" name="sms_time" size="4" value="<%=rs("sms_time")%>" <%=onKeyUp%>> Сʱ��(0Ϊ������)</td>
                </tr>
                 <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">�Զ�ɾ����֤��¼ʱ�䣺</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="sms_del" size="5" value="<%=rs("sms_del")%>" <%=onKeyUp%>> �� ��0Ϊ��ɾ�����������ó���30���Զ�ɾ����ɾ���󽫲���������ע������ֻ���IP��ַ��������Ҫ���á���</td>
                </tr>
                <tr bgcolor="#BDBEDE">
                  <td height="28" colspan="2" style="border-top-style: none; border-top-width: medium; border-bottom-style: solid; border-bottom-width: 1">
                    <span class="style1"><font size="2">&nbsp;VIP�趨</font></span></td>
                </tr>
                 <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">VIP�ʸ���Ч���������룺</span></td>
                  <td style="border-bottom-style: none; border-bottom-width: medium" height="29" bgcolor="#FFFFFF"><p style="margin-left: 20" >
                      <input type="text" name="vipsj" value="<%=rs("vipsj")%>" size="20"<%=onKeyUp%> />
���� (����ʱ�䳤�㣬���ڹ���)</td> 
                </tr>               
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">VIP��Ա�ʸ���Ҫ��</span></td>
                  <td style="border-bottom-style: none; border-bottom-width: medium" height="29" bgcolor="#FFFFFF"><p style="margin-left: 20">
                      <input type="text" name="vipje" value="<%=rs("vipje")%>" size="20"  <%=onKeyUp%>/>
Ԫ�����/�� (����Ϊ����0������)</td>
                </tr>

                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">��VIP�û�ֻ��ʾ��</span></td>
                  <td style="border-bottom-style: none; border-bottom-width: medium" height="29" bgcolor="#FFFFFF"><p style="margin-left: 20">
            <input type="text" name="huiyuansp" value="<%=rs("huiyuansp")%>" size="20"  <%=onKeyUp%>/>
��������Ʒ</td>
                </tr>
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">��VIP�û�ֻ��ʾ��</span></td>
                  
				  
				  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="huiyuanxx" type="text" id="huiyuanxx" value="<%=rs("huiyuanxx")%>" size="20" <%=onKeyUp%> />
				    ��������Ϣ</p></td>
                </tr>
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">�����������Ҫ������</span></td>
				  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="xinxis" type="text" id="xinxis" value="<%=rs("xinxis")%>" size="20" <%=onKeyUp%> />
				    ��������Ϣ</p></td>
                </tr>
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">VIP��Ա�������·����ޣ�</span></td>
				  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="usernews" type="text" id="usernews" value="<%=rs("usernews")%>" size="20" <%=onKeyUp%> />
				    ƪ��0�����ƣ�</p></td>
                 </tr>
                 <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">VIP��Ա�������������ޣ�</span></td>
				  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="userlink" type="text" id="userlink" value="<%=rs("userlink")%>" size="20" <%=onKeyUp%> />
				    ����0�����ƣ�</p></td>
                </tr>
                 <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">VIP�������ι��������ޣ�</span></td>
				  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="adjf" type="text" id="adjf" value="<%=rs("adjf")%>" size="20" <%=onKeyUp%> />
				    �֣���������ʾ��վ��棬0ʼ����ʾ��</p></td>
                </tr>
<%
certificate=rs("certificate")
certificate=split(certificate,"|")
certificate1=certificate(0)
certificate2=certificate(1)
certificate3=certificate(2)
certificate4=certificate(3)
certificate5=certificate(4)
certificate6=certificate(5)
certificate7=certificate(6)
%>
                 <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">���������ƬҪ��˵�֤�գ�</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20"><INPUT TYPE=checkbox NAME="certificate1" VALUE="1" <%if certificate1="1" then response.write("checked")%>>�������֤&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="certificate2" VALUE="1" <%if certificate2="1" then response.write("checked")%>>���˱�׼��&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="certificate3" VALUE="1" <%if certificate3="1" then response.write("checked")%>>Ӫҵִ��&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="certificate4" VALUE="1" <%if certificate4="1" then response.write("checked")%>>��˰�Ǽ�֤&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="certificate5" VALUE="1" <%if certificate5="1" then response.write("checked")%>>��˰�Ǽ�֤<br><INPUT TYPE=checkbox NAME="certificate6" VALUE="1" <%if certificate6="1" then response.write("checked")%>>��֯��������֤
					&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="certificate7" VALUE="1" <%if certificate7="1" then response.write("checked")%>>����֤��(���桢���䡢�̱��)
&nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>��Ա�������ʱ�ɿ������������</span></a></td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">��������������ѣ�</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><textarea name="prompt" cols="70" rows="3" class="input2" onpropertychange="kn()"><%=rs("prompt")%></textarea>
                  &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>���ѻ�Ա�������Ӧע�����������200�ַ��ڡ�</span></a><br><span id=num>������������200���ֽڣ�100����</span></td>
                </tr>
                <tr bgcolor="#BDBEDE">
                  <td height="28" colspan="2" bgcolor="#BDBEDE" style="border-top-style: none; border-top-width: medium; border-bottom-style: solid; border-bottom-width: 1"><span class="style1"><font size="2">&nbsp;������Ϣ��������</font></span></td>
                </tr>

                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">�Ƿ�ע��ſɷ�����</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="addxinxi" type="text" value="<%=rs("addxinxi")%>" size="20" maxlength="1" <%=onKeyUp%> />
                    0=ע���Ա���ܷ�����1=����ע��ֱ�ӷ���</p></td>
                </tr>
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">ע����ò��ܷ�����</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="zcfbsj" type="text" value="<%=rs("zcfbsj")%>" size="20" maxlength="2" <%=onKeyUp%> />
                    ��&nbsp;&nbsp;Ϊ��ֹȺ��������Ϣ������60������</p></td>
                </tr>
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right"><span class="style2" style="margin-left: 20">�οͷ�����Ϣ���ã�</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="xinxish" type="text" value="<%=rs("xinxish")%>" size="20" maxlength="1" <%=onKeyUp%> />
                    0=�ȴ���˲���ͨ����1=�������ֱ��ͨ�� <font color="#ff0000">����ʹ��</font></p></td>
                </tr>
                 <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right"><span class="style2" style="margin-left: 20">��VIP������Ϣ���ã�</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="vipno" type="text" value="<%=rs("vipno")%>" size="20" maxlength="1" <%=onKeyUp%> />
                    0=�ȴ���˲���ͨ����1=�������ֱ��ͨ�� <font color="#ff0000">����ʹ��</font></p></td>
                </tr>
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right"><span class="style2" style="margin-left: 20">������Ϣ����ܿ��أ�</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="onoff" type="text" value="<%=rs("onoff")%>" size="20" maxlength="1" <%=onKeyUp%> />
                    0=�ȴ���˲���ͨ����1=���Ϸ���������ͨ�� <font color="#ff0000">����ʹ��</font></p></td>
                </tr>
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">��Ϣ���ⳤ�����ƣ�</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="biaotinum" type="text" value="<%=rs("biaotinum")%>" size="20" <%=onKeyUp%>  />
                    ���ֽڣ�һ������2���ֽڣ�������40�ֽ���</p></td>
                </tr>
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">��Ϣ˵���������ƣ�</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="memonum" type="text" value="<%=rs("memonum")%>" size="20" <%=onKeyUp%>  />
                    ���ֽڣ�һ������2���ֽڣ�������2000�ֽ���</p></td>
              </tr>
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">������Ϣ�۳����֣�</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="gxkf" type="text" value="<%=rs("gxkf")%>" size="20" <%=onKeyUp%>  />
                    �� ��Աһ��������Ϣ�����º����п�ǰ�����۳��Ļ��֣�0Ϊ���۳�</p></td>
                </tr>			
                <tr bgcolor="#BDBEDE">
                  <td height="28" colspan="2" style="border-top-style: none; border-top-width: medium; border-bottom-style: solid; border-bottom-width: 1">
                    <span class="style1"><font size="2">&nbsp;��Ϣ��ʾ����</font></span></td>
                </tr>
                <!--<tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">�Ƿ����ɾ�̬��ҳ��</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="HtmlHome" type="text" value="<%=rs("HtmlHome")%>" size="20" maxlength="1" <%=onKeyUp%> />
                      <label></label>
                    0=asp��̬��ҳ��ʽ��ʾ��1=HTM��̬��ҳ��ʽ��ʾ<br><font color="#ff0000">(���ѡ�񰴵�������վ����ʽ��ʾ�й���������ѡ��̬��ҳ��ԭ���ǲ���̬�����Զ���������ʾ��Ϣ��)</font></p></td>
                </tr>-->
				<input type="hidden" name="HtmlHome" value="0">
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">�Ƿ����ɾ�̬����ҳ��</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="asphtm" type="text" value="<%=rs("asphtm")%>" size="20" maxlength="1" <%=onKeyUp%> />
                      <label></label>
                    0=asp��̬ҳ��ʽ��ʾ��1=HTM��̬ҳ��ʽ��ʾ<br><font color="#ff0000">(��̬����ҳ���������ٶȡ�����������¼�ͼ��������ѹ������ռ�ÿռ䣬�û�����������)</font></p></td>
                </tr>
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">����ǰ̨������Ϣ����HTM���룺</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="Filterhtm" type="text" value="<%=rs("Filterhtm")%>" size="20" <%=onKeyUp%>  />
                    0�����ˣ�1�����οͣ�2���˷�VIP��Ա��3ȫ������</p></td>
                </tr>
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">��Ϣ��ϸ������ʾ���ȣ�</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="content_length" type="text" value="<%=rs("content_length")%>" size="20" <%=onKeyUp%> />
                      <label></label>
                    ���ֽ�  0Ϊ����ʾ���ݣ��������趨�ֽڳ��ȣ�һ������Ϊ2���ֽ�<br><font color="#ff0000">(��������Ϣѡ���Ķ�Ȩ��ʱ������)</font></p></td>
                </tr>
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">��ɽ����Ƿ���ʾ��</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="jiaoyi" type="text" value="<%=rs("jiaoyi")%>" size="20" maxlength="1" <%=onKeyUp%> />
                      <label></label>
                    0=����ʾ�ѽ�����ɣ�1=ͬʱ��ʾ�������</p></td>
                </tr>
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">������Ϣ�Ƿ���ʾ��</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="overdue" type="text" value="<%=rs("overdue")%>" size="20" maxlength="1" <%=onKeyUp%> />
                      <label></label>
                    0=����ʾ������Ϣ��1=ͬʱ��ʾ������Ϣ</p></td>
                </tr>

                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">ʹ���ö�������Ч�ڣ�</td>
                 <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="b_y" size="20" value="<%=rs("b_y")%>" <%=onKeyUp%>> 
��&nbsp;&nbsp;ϵͳ��ȡ��������������Ϣ�ö�(0Ϊ�������ƣ�<font color="#ff0000">����ʹ��</font>)</td>
                </tr>
                  <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">�Ƽ�ʱ����Ч�ڣ�</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="tui_y" size="20" value="<%=rs("tui_y")%>" <%=onKeyUp%>> 
��&nbsp;&nbsp;ϵͳ��ȡ��������������Ϣ�Ƽ�(0Ϊ�������ƣ�<font color="#ff0000">����ʹ��</font>)</td>
                </tr>
                  <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">������Ϣ��׼��</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="hotbz" size="20" value="<%=rs("hotbz")%>" <%=onKeyUp%>> 
��&nbsp;&nbsp;������Ϣ��Ҫ�ﵽ�ĵ����</td>
                </tr>
<%
HomeElite=rs("HomeElite")
HomeElite=split(HomeElite,"|")
HomeElite1=HomeElite(0)
HomeElite2=HomeElite(1)
HomeElite3=HomeElite(2)
HomeElite4=HomeElite(3)
HomeElite5=HomeElite(4)
%>
                  <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">��ҳͷ��ͼƬ���������</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20"><INPUT type="text" NAME="HomeElite1" size="5" VALUE="<%=HomeElite1%>" <%=onKeyUp%>>ͼƬ��� &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <INPUT type="text" NAME="HomeElite2" size="5" VALUE="<%=HomeElite2%>" <%=onKeyUp%>>ͼƬ�߶� &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT type="text" NAME="HomeElite3" size="5" VALUE="<%=HomeElite3%>" <%=onKeyUp%>>�����С &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <br><INPUT type="text" NAME="HomeElite4" size="15" VALUE="<%=HomeElite4%>">�� ��<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>�������ϱ����д����壬��������ʱ����Ĭ�ϵ����壡</span></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT type="text" NAME="HomeElite5" size="15" VALUE="<%=HomeElite5%>">�洢λ��<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>ͼƬ�洢λ������UploadFiles/Home/�������ԡ�/�����ߡ�..����ͷ��Ŀ¼Ҫ�����ֹ�������</span></a>&nbsp;&nbsp;<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>����ҳ��Ӧλ�ù�񲻸ı������޸�ͷ��ͼƬ���������</span></a></td>
                </tr> 


				            							
<%
codekg=rs("codekg")
codekg=split(codekg,"|")
codekg1=codekg(0)
codekg2=codekg(1)
codekg3=codekg(2)
codekg4=codekg(3)
codekg5=codekg(4)
%>
                  <tr bgcolor="#BDBEDE">
                  <td height="28" colspan="2" style="border-top-style: none; border-top-width: medium; border-bottom-style: solid; border-bottom-width: 1">
                    <span class="style1"><font size="2">&nbsp;��֤��������</font></span></td>
                </tr>
				<tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">ѡ����֤������</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20"><INPUT TYPE=checkbox NAME="codekg1" VALUE="1" <%if codekg1="1" then response.write("checked")%>>�ʴ���֤&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="codekg2" VALUE="1" <%if codekg2="1" then response.write("checked")%>>������֤&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="codekg3" VALUE="1" <%if codekg3="1" then response.write("checked")%>>������֤&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="codekg4" VALUE="1" <%if codekg4="1" then response.write("checked")%>>������֤
					&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="codekg5" VALUE="1" <%if codekg5="1" then response.write("checked")%>>��ʽ��֤
&nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>�����ַ�ʽ��֤���ɶ�ѡ�;����任����ˮ����������ѡ���ʴ���֤�룡</span></a></td>
                </tr>
                <tr bgcolor="#BDBEDE">
                  <td height="28" colspan="2" style="border-top-style: none; border-top-width: medium; border-bottom-style: solid; border-bottom-width: 1">
                    <span class="style1"><font size="2">&nbsp;��������</font></span></td>
                </tr>
                  <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">�������ർ����ʾ������</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="citys" size="20" value="<%=rs("citys")%>" <%=onKeyUp%>> 
��&nbsp;&nbsp;�ɸ��ݷ������Ƴ����ʵ�����</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">��ҳ�õƹ����ʾ���</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="radio" name="slidelx" value="0" <%if rs("slidelx")=0 then%>checked<%end if%>>�õƹ��A���ֵ�����ʾ�� <input type="radio" name="slidelx" value="1" <%if rs("slidelx")=1 then%>checked<%end if%>>�õƹ��B������ʾ��վ��棩 </td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">��ҳ��Ƶ��ʾ���</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="radio" name="hdlb" value="1" <%if rs("hdlb")=1 then%>checked<%end if%>>��ͨ��Ƶ��� <input type="radio" name="hdlb" value="2" <%if rs("hdlb")=2 then%>checked<%end if%>>FLV��Ƶ��� <input type="radio" name="hdlb" value="0" <%if rs("hdlb")=0 then%>checked<%end if%>>�ر���Ƶ���(�����ű�׼���ID:240)</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">������ѶͶ�忪�أ�</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="radio" name="contribute" value="1" <%if rs("contribute")=1 then%>checked<%end if%>>�ο� <input type="radio" name="contribute" value="2" <%if rs("contribute")=2 then%>checked<%end if%>>��ͨ��Ա <input type="radio" name="contribute" value="3" <%if rs("contribute")=3 then%>checked<%end if%>>VIP��Ա <input type="radio" name="contribute" value="0" <%if rs("contribute")=0 then%>checked<%end if%>>�ر�Ͷ����� ����������¼��ݣ�</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">��Ϣ�ظ����������ۿ��أ�</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="Newspl" value="<%=rs("Newspl")%>" size="20">
                  &nbsp;&nbsp; 0Ϊȫ���ţ�1Ϊ��Ҫ��¼�������ۣ�2Ϊ��ʾ���۽�ֹ�����ۣ�3Ϊȫ����ֹ</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">��Ϣ�ظ�������������ˣ�</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="plsh" value="<%=rs("plsh")%>" size="20">
                  &nbsp;&nbsp; 0Ϊֱ����ʾ��1Ϊ�����ʾ</td>
                </tr>				
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">��վ�������ƣ�</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="CacheName" value="<%=rs("CacheName")%>" size="20">
                  &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>������һ�������ո��Ӣ�����Ƽ��ɣ�����ʹ�ú��֣�����ʹ��ϵͳĬ�ϵġ�</span></a></td>
                </tr>
			
              </table>
            </td>
          </tr>
 

          <tr>
            <td align="center" width="688" bgcolor="#eeeeee" height="35">            <table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><div align="center">
                  <input type="submit" value="ȷ������" name="B1">
                </div></td>
                <td><div align="center">
                  <input type="reset" value="ȡ������" name="B1">
                </div></td>
              </tr>
            </table></td>
          </tr>
        </table>
<%
    end if
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
%>
      </form>
    </td>
  </tr>
</table>
