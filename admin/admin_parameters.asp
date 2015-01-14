<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%If (request("shows1")=1 Or request("shows2")=1 Or request("shows3")=1 Or request("shows4")=1 Or request("shows5")=1 Or request("shows6")=1 Or request("shows7")=1 Or request("shows8")=1 Or request("shows9")=1 Or request("shows10")=1) And request("HtmlHome")=1 Then
Response.Write ("<script language=javascript>alert('选择按地区（分站）方式显示有关内容时不宜生成静态首页!');history.go(-1);</script>")
Response.end
End If
If request("ok")=1 And request("leixing")="" Then
Response.Write ("<script language=javascript>alert('信息类型不能为空，且要注意格式规则!');history.go(-1);</script>")
Response.end
End If
If request("ok")=1 And (request("S10")="" or request("S11")="") Then
Response.Write ("<script language=javascript>alert('QQ客服号码和昵称不能为空，且要注意格式!');history.go(-1);</script>")
Response.end
End If
If request("mailqr")=1 And request("regmail")=0 Then
Response.Write ("<script language=javascript>alert('选择邮件确认时必须同时选择邮件通知!');history.go(-1);</script>")
Response.end
End If
If request("ok")=1 And request("vipje")<1  Then
Response.Write ("<script language=javascript>alert('[VIP会员资格需要]选项为必须大于0的整数!');history.go(-1);</script>")
Response.end
End If


dim ObjInstalled
Const Script_FSO="Scripting.FileSystemObject"
ObjInstalled=IsObjInstalled(Script_FSO)
If ObjInstalled=false Then
 Response.Write "<b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能。 请直接修改“根目录下的config.asp和Database.asp”文件中的内容。</font></b>"
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
 hf.write "'=====以下参数必须保持正确，由后台设置生成，可用记事本手工修改，不明白就不要乱改动====www.ip126.com===" & vbcrlf 
 hf.write vbcrlf
 hf.write "'==========================数据库类型====================================" & vbcrlf
 hf.write "Const SystemDatabaseType=" & trim(request("SystemDatabaseType")) & " '数据库类型 ACCESS数据库为0 SQL数据库为1 " & vbcrlf
 hf.write vbcrlf
 hf.write "'==========================Access数据库==================================" & vbcrlf
 hf.write "Const DBFileName=" & chr(34) & trim(request("DBFileName")) & chr(34) & " 'ACCESS数据库存放的路径, 相对于根目录的路径" & vbcrlf
 hf.write vbcrlf
 hf.write "'==========================SQL数据库=====================================" & vbcrlf
 hf.write "Const SqlDatabaseName=" & chr(34) & trim(request("SqlDatabaseName")) & chr(34) & " 'SQL数据库名(SqlDatabaseName)" & vbcrlf
 hf.write "Const SqlUsername=" & chr(34) & trim(request("SqlUsername")) & chr(34) & " 'SQL数据库登录用户名(SqlUsername)" & vbcrlf
 hf.write "Const SqlPassword=" & chr(34) & trim(request("SqlPassword")) & chr(34) & "'SQL数据库用户登录密码(SqlPassword)" & vbcrlf
 hf.write "Const SqlHostIP=" & chr(34) & trim(request("SqlHostIP")) & chr(34) & "" & vbcrlf
 hf.write "'SQL主机地址(SqlHostIP)设置注意：程序与数据库同机用(local)；远程用IP，如""219.48.165.56""；数据库地址为域名的用""数据库域名""；本地调试时若与SQL Server其它版本并存安装时，SQL主机地址要这样填写：""计算机名\实例名""，如""user\sql2008""" & vbcrlf
 
 hf.write "%" & ">" & vbcrlf
 hf.close
 set hf=nothing
 set fso=Nothing
'================当设置首页动态页时删除已生成的静态首页文件
If request("HtmlHome")=0 Then
Dim objFSO
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if objFSO.fileExists(Server.MapPath("../index..html")) then
   objFSO.DeleteFile(Server.MapPath("../index..html"))
end if
set objfso=Nothing
End if
'===============================生成客服代码
Dim JsCode,JsFileName
JsCode=Html2Js(request("k54kefu"))
JsCode="document.write('"&JsCode&"')"
JsCode = "<!--" & vbCrLf & JsCode & vbCrLf &"//-->" 
JsFileName = Server.MapPath("..\UploadFiles\js\54kefu.js")
Call WriteToFile(JsFileName,JsCode,"utf-8")'用函数方式

url="admin_parameters.asp?id=1"
response.Write "<script language=javascript>alert('操作成功！将自动生成配置文件config.asp');location.href='admin_config.asp?page="&url&"';</script>"
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
/*提示信息*/
.info {position:relative;background:#;color:#666; text-decoration:none;font-size:12px;width:0px;height:0px;text-align:center;border:0px solid #ccc;line-height:18px;}/*设置链接的属性,一定要设置为relative才能使提示层跟着链接走*/
.info:hover {background:#eee;color:#333;}
.info span {display: none }/*设置正常下的span为隐藏状态*/
.info:hover span /*设置hover下的span属性为呈现状态,并设置提示层的位置*/{display:block;position:absolute;top:0px;left:25px;width:200px;
border:1px solid #DC5826; background:#FFF3B4; color:#333;padding:5px;text-align:left;}
</style>
<script language="javascript">
<!--
//说明字节数限制，与下面配合=========
String.prototype.Length=function(){
temp=this.replace(/([^\x00-\xff])/g,"$1$1");
return temp.length
}
//显示输入字符数
function kn()
  {
   var ln=myform.prompt.value.Length();
     num.innerText='您还可以输入:'+(200-ln)+'个字符'+(<%=200/2%>-ln/2)+'汉字';
      //if (ln>=8) form.memo.readOnly=true;  //这行是如果满足条件表单无法输入和修改
}

function myform_onsubmit() {
if (document.myform.S0.value.length>50)
	{
	  alert("网站名称不能超过50个字符")
	  document.myform.S0.focus()
	  return false
	 }
if (document.myform.web.value.length>50)
	{
	  alert("网站域名不能超过50个字符")
	  document.myform.web.focus()
	  return false
	 }

function contain(str,charset)//字符串包含测试函数
{
//下面三行是字串中不能包含指定字符中的任何一个
//var i;
//for(i=0;i<charset.length;i++)
//if(str.indexOf(charset.charAt(i))>=0)
if(str.indexOf(charset)>=0)//此行为字串中不能包含指定字符的整体
return true;
return false;
} 
if(contain(document.myform.web.value,"http://"))
{ 
alert("网址前不要带http://");
document.myform.web.focus();
return false;
}
if(contain(document.myform.web.value,"/"))
{ 
alert("网址后不要带/");
document.myform.web.focus();
return false;
}

//验证邮箱正确性
function checkemail(email){
var str=email;
//在JavaScript中，正则表达式只能使用"/"开头和结束，不能使用双引号
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
    alert("您输入Email地址不正确，请重新输入!");
    document.myform.S4.focus();
    return false;
    }

if (document.myform.prompt.value.length>200)
	{
	  alert("店铺申请条件提醒内容不能超过200个字符")
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
            <p class="style1"><font size="2">&nbsp;基本参数设定</font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" align="center" style="border-left-style: none; border-left-width: medium; border-right-style: none; border-right-width: medium" width="700" height="735"> 
              <table width="100%" border="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
              
                <tr> 
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium; border-top-style:solid; border-top-width:1" height="27" bgcolor="#FFFFFF" align="right">设定网站名称：</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium; border-top-style:solid; border-top-width:1" height="27" bgcolor="#FFFFFF"> 
                    <p style="margin-left: 20"> 
                    <input type="text" name="S0" value="<%=Application(CacheName&"_WebSetting")(0)%>" size="40"> 
                    </td>
                </tr>
                <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium; border-top-style:none; border-top-width:medium" height="27" bgcolor="#FFFFFF" align="right">网站域名：</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium; border-top-style:none; border-top-width:medium" height="27" bgcolor="#FFFFFF"> 
                    <p style="margin-left: 20"> 
                    <input type="text" name="web" value="<%=rs("web")%>" size="40">(前面不带“http://”，后面不带“/”，若在根目录的下级目录安装本系统，后面不要带该目录名)
                    </td>
                </tr>
                <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium; border-top-style:none; border-top-width:medium" height="27" bgcolor="#FFFFFF" align="right"><font color="#FF0000">系统安装目录(重要)：</font></td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium; border-top-style:none; border-top-width:medium" height="27" bgcolor="#FFFFFF"> 
                    <p style="margin-left: 20"> 
                    <input type="text" name="InstallDir" value="<%=strInstallDir%>" size="40"> &nbsp;<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>(相对于根目录的位置，前后都不带“/”,若在根目录请留空。目录更改后，要重新生成所有模板、htm文件和广告图片路径、代码等)</span></a></td>
                </tr>
				 <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium; border-top-style:none; border-top-width:medium" height="27" bgcolor="#FFFFFF" align="right">数据库类型：</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium; border-top-style:none; border-top-width:medium"  bgcolor="#FFFFFF"><p style="margin-left: 20"><input name="SystemDatabaseType" onClick="javascript:changedbtype(0);" type="radio" value="0" <%if SystemDatabaseType=0 then%> checked<%end if%>>ACCESS
<input type="radio" onClick="javascript:changedbtype('1');" name="SystemDatabaseType" value="1" <%if SystemDatabaseType="1" then%> checked<%end if%>>SQL（仅限企业版）采用SQL存储过程会有更好的性能表现<br><font color="#FF0000">根据版本正确选定，不能乱设，否则无法访问数据库！</font></td>
                </tr>

<tr id="accesstr" name="accesstr"> 
<td width="30%" style="border-bottom-style: none; border-bottom-width: medium; border-top-style:none; border-top-width:medium" height="27" bgcolor="#FFFFFF" align="right">ACCSS数据库文件路径：</td>
<td><p style="margin-left: 20"><input name="DBFileName" type="text" id="DBFileName" value="<%=DBFileName%>" size="40" maxlength="50"> <br>扩展名改为asp可防下载。更名后要手工修改相应的目录和数据库文件名</td>
</tr>
<tr id="sqltr" name="sqltr"> 
<td width="30%" style="border-bottom-style: none; border-bottom-width: medium; border-top-style:none; border-top-width:medium" height="27" bgcolor="#FFFFFF" align="right">MS SQL数据库设置：</td>
<td colspan="2"><p style="margin-left: 20"><table  height="80" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td><input name="SqlDatabaseName" type="text" id="SqlDatabaseName" value="<%=SqlDatabaseName%>" size="25" maxlength="50"> SQL数据库名</td>
  </tr>
  <tr>
    <td><input name="SqlUsername" type="password" id="SqlUsername" value="<%=SqlUsername%>" size="25" maxlength="50"> 登录数据库用户名</td>
  </tr>
  <tr>
    <td><input name="SqlPassword" type="password" id="SqlPassword" value="<%=SqlPassword%>" size="25" maxlength="50"> 登录数据库密码</td>
  </tr>
  <tr>
    <td><input name="SqlHostIP" type="text" id="SqlLocalName" value="<%=SqlHostIP%>" size="25" maxlength="50"> 连接数据源[站库同机用(local)，站库异机用IP或数据源域名]</td>
  </tr>
</table></td>
</tr>
<script language="javascript">changedbtype(<%=SystemDatabaseType%>);</script>				
                <tr> 
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">网站友情链接LOGO：</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="logo" value="<%=rs("logo")%>" size="40"> 
                    （大小88*31，gif格式）</td>
                </tr>
                <tr> 
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">网站经营备案号：</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="icp" value="<%=rs("icp")%>" size="40"> 
                    </td>
                </tr>
                 <tr> 
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">公司名称：</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="S12" value="<%=Application(CacheName&"_WebSetting")(12)%>" size="40"> 
                    </td>
                </tr>
                 <tr> 
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">办公地址：</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="S8" value="<%=Application(CacheName&"_WebSetting")(8)%>" size="40"> 
                    </td>
                </tr>
                 <tr> 
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">邮政编码：</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="S9" value="<%=Application(CacheName&"_WebSetting")(9)%>" size="30" <%=onKeyUp%>> 
                    </td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">办公电话：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="S5" value="<%=Application(CacheName&"_WebSetting")(5)%>" size="30">
                  </td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">联系手机：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="S6" value="<%=Application(CacheName&"_WebSetting")(6)%>" size="30" <%=onKeyUp%>>
                  </td>
                </tr>
                 <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">传真号码：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="S7" value="<%=Application(CacheName&"_WebSetting")(7)%>" size="30">
                  </td>
                </tr>
                  <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">电子信箱：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="S4" value="<%=Application(CacheName&"_WebSetting")(4)%>" size="30"> (可与邮件系统管理中的发信人信箱相同)
                  </td>
                </tr>               
 

                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">网站停止运行说明：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><textarea name="stopsm" cols="80" rows="3" class="input2"><%=rs("stopsm")%></textarea></td>   
                </tr>
           
                    
                <tr> 
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF" align="right">交易类型设置：</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF">
                    <p style="margin-left: 20"><input type="text" name="leixing" value="<%=rs("leixing")%>" size="45"> <br><font color="#FF0000">
类型之间用半角的“|”分隔，前后不带“|”。例</font>：类别一|类别二|类别三                                      
                    </td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">网站标题栏信息：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><textarea name="S1" cols="80" rows="3" class="input2"><%=Application(CacheName&"_WebSetting")(1)%></textarea></td>
                </tr>
                  <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">网站关键词(keywords)：<br>供搜索引擎搜索的关键内容，关键字用“,”号分隔</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><textarea name="S3" cols="80" rows="5" class="input2"><%=Application(CacheName&"_WebSetting")(3)%></textarea></td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">网站说明(description)：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><textarea name="S2" cols="80" rows="5" class="input2"><%=Application(CacheName&"_WebSetting")(2)%></textarea></td>
                </tr>

                 <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">网站简介：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><textarea name="synopis" cols="80" rows="5" class="input2"><%=rs("synopis")%></textarea></td>
                </tr>
                 <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">统计代码：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><textarea name="tjdm" cols="80" rows="5" class="input2"><%=rs("tjdm")%></textarea> 到有关流量统计服务商处申请一个流量统计代码填在此并生成尾部模板就可统计了</td>
                </tr>

                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">友情链接申请：
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"> <%if rs("addlink")=0 then%>               
                <input type="radio" name="addlink" value="0" id="addlink" checked>
                 允许&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="addlink" value="1" id="addlink">
                禁止
                <%else%>                         
                <input type="radio" name="addlink" value="0" id="addlink">
                 允许&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="addlink" value="1" id="addlink" checked>
                禁止<%end if%> 
                  </td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">友情链接上传LOGO：
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"> <%if rs("linklogo")=1 then%>               
                <input type="radio" name="linklogo" value="1" id="linklogo" checked>
                 允许&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="linklogo" value="0" id="linklogo">
                禁止
                <%else%>                         
                <input type="radio" name="linklogo" value="1" id="linklogo">
                 允许&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="linklogo" value="0" id="linklogo" checked>
                禁止<%end if%> 
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
                    <span class="style1"><font size="2">&nbsp;首页板块开关设置</font></span></td>
                </tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">首页板块开关：</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20"><INPUT TYPE=checkbox NAME="lmkg1" VALUE="1" <%if lmkg1="1" then response.write("checked")%>>开放网站&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="lmkg2" VALUE="1" <%if lmkg2="1" then response.write("checked")%>>网站只读&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="lmkg3" VALUE="1" <%if lmkg3="1" then response.write("checked")%>>城市导航&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="lmkg4" VALUE="1" <%if lmkg4="1" then response.write("checked")%>>栏目列表
					&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="lmkg5" VALUE="1" <%if lmkg5="1" then response.write("checked")%>>竞价信息
					&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="lmkg6" VALUE="1" <%if lmkg6="1" then response.write("checked")%>>企业名片<br>
					<INPUT TYPE=checkbox NAME="lmkg7" VALUE="1" <%if lmkg7="1" then response.write("checked")%>>推荐店铺
					&nbsp;<INPUT TYPE=checkbox NAME="lmkg8" VALUE="1" <%if lmkg8="1" then response.write("checked")%>>都市114					
					&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="lmkg9" VALUE="1" <%if lmkg9="1" then response.write("checked")%>>翻牌广告
					&nbsp;<INPUT TYPE=checkbox NAME="lmkg10" VALUE="1" <%if lmkg10="1" then response.write("checked")%>>友情链接&nbsp;&nbsp; <INPUT TYPE=checkbox NAME="lmkg11" VALUE="1" <%if lmkg11="1" then response.write("checked")%>>图片信息
					&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="lmkg12" VALUE="1" <%if lmkg12="1" then response.write("checked")%>>便民服务<br>
					<INPUT TYPE=checkbox NAME="lmkg13" VALUE="1" <%if lmkg13="1" then response.write("checked")%>>资源共享&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="lmkg14" VALUE="1" <%if lmkg14="1" then response.write("checked")%>>新闻相册&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="lmkg15" VALUE="1" <%if lmkg15="1" then response.write("checked")%>>电子杂志
<br>可根据需要开关首页版块</td>
                </tr>
                  <tr bgcolor="#BDBEDE">
                  <td height="28" colspan="2" style="border-top-style: none; border-top-width: medium; border-bottom-style: solid; border-bottom-width: 1">
                    <span class="style1"><font size="2">&nbsp;地区显示设置</font></span></td>
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
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">自动进入本地：
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20">
		 <%Dim rscip
		  set rscip=server.CreateObject("adodb.recordset")
		  rscip.Open "select * from DNJY_city where id>0 and twoid=0 order by id",conn,1,1		  
		  if rscip.EOF and rscip.BOF then%>	
          <input type="hidden" name="cip" value="0" id="cip">
			（地区分类目前为空，启用无效）
         <%else%>
				<%if rs("cip")=True then%>               
                <input type="radio" name="cip" value="1" id="cip" checked>
                 启用&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="cip" value="0" id="cip">
                关闭
                <%else%>                         
                <input type="radio" name="cip" value="1" id="cip">
                 启用&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="cip" value="0" id="cip" checked>
                关闭<%end if%> 
		  <%rscip.close
          set rscip=Nothing
          End If%>
                  <span class="style2"><br>选择启用时自动根据访客IP地址进入所属本地分站（IP地址资料复杂可能有不准确现象，且如果地区分类太多，如全国省市县三级，网站读取数据和比对要一定时间，显得网站速度有迟滞现象，谨慎启用！建议根据地区分类数量开关对比速度确定是否启用。）</span></td>
                </tr>
				<tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">板块分站显示：</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20"><INPUT TYPE=checkbox NAME="shows1" VALUE="1" <%if shows1="1" then response.write("checked")%>>供求信息&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="shows2" VALUE="1" <%if shows2="1" then response.write("checked")%>>新闻资讯&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="shows3" VALUE="1" <%if shows3="1" then response.write("checked")%>>商家店铺&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="shows4" VALUE="1" <%if shows4="1" then response.write("checked")%>>企业名片
					&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="shows5" VALUE="1" <%if shows5="1" then response.write("checked")%>>用户留言&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="shows6" VALUE="1" <%if shows6="1" then response.write("checked")%>>都市114 <br>
					<INPUT TYPE=checkbox NAME="shows7" VALUE="1" <%if shows7="1" then response.write("checked")%>>便民信息
					&nbsp;<INPUT TYPE=checkbox NAME="shows8" VALUE="1" <%if shows8="1" then response.write("checked")%>>网站广告&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="shows9" VALUE="1" <%if shows9="1" then response.write("checked")%>>友情链接&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="shows10" VALUE="1" <%if shows10="1" then response.write("checked")%>>网站LOGO&nbsp;&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="shows11" VALUE="1" <%if shows11="1" then response.write("checked")%>>新闻相册&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="shows12" VALUE="1" <%if shows12="1" then response.write("checked")%>>电子杂志
					&nbsp;
<br><span class="style2">如果不选择则该板块按总站（全地区）内容显示</span></td>
                </tr> 				
                <tr bgcolor="#BDBEDE">
                  <td height="28" colspan="2" style="border-top-style: none; border-top-width: medium; border-bottom-style: solid; border-bottom-width: 1">
                    <span class="style1"><font size="2">&nbsp;上传图片水印设置</font></span></td>
                </tr>
               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right"></td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20">＝＝＝＝＝＝＝＝＝＝以下为信息和新闻资讯上传图片水印设置＝＝＝＝＝＝＝＝＝＝＝<br>
				  （AspJpeg组件支持：<%if IsObjInstalled("Persits.Jpeg") then%>√<%else%>×<%end if%>&nbsp;<span class="fontcolor">服务器必须支持AspJpeg组件，否则系统将自动关闭本功能）</span></td>
                </tr>
               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">水印类型：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="radio" name="SY_Type" value="0" <%if rs("SY_Type")=0 then%>checked<%end if%>>不加水印	<input type="radio" name="SY_Type" value="1" <%if rs("SY_Type")=1 then%>checked<%end if%>>文字水印	<input type="radio" name="SY_Type" value="2" <%if rs("SY_Type")=2 then%>checked<%end if%>>图片水印	</td>
                </tr>
               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">限制时间：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="SY_interval" value="<%=rs("SY_interval")%>" size="10" <%=onKeyUp%>>
                  分钟内编辑的图片不重复加水印</td>
                </tr>
               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">水印位置坐标：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20">X：<input name='SY_X' type='text' value='<%=rs("SY_X")%>' size='10' maxlength='10'  ONKEYPRESS="javascript:event.returnValue=IsDigit()" <%=onKeyUp%>> 像素<br>Y：<input name='SY_Y' type='text' value='<%=rs("SY_Y")%>' size='10' maxlength='10'  ONKEYPRESS="javascript:event.returnValue=IsDigit()" <%=onKeyUp%>> 像素</td>
                </tr>
               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">文字水印：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="SY_Text" value="<%=rs("SY_Text")%>" size="30">
                  文字字数不宜超过15个字符，不支持任何WEB编码标记</td>
                </tr>
               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">文字字体：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20">
			<SELECT name="SY_FontName" >
            <option value="宋体"  <% if rs("SY_FontName")="宋体" Then Response.write("Selected") %>>宋体</option>
            <option value="楷体_GB2312" <% if rs("SY_FontName")="楷体_GB2312" Then Response.write("Selected") %>>楷体</option>
            <option value="仿宋_GB2312" <% if rs("SY_FontName")="仿宋_GB2312" Then Response.write("Selected") %>>仿宋体</option>
            <option value="黑体" <% if rs("SY_FontName")="黑体" Then Response.write("Selected") %>>黑体</option>
            <option value="隶书" <% if rs("SY_FontName")="隶书" Then Response.write("Selected") %>>隶书</option>
            <option value="幼圆" <% if rs("SY_FontName")="幼圆" Then Response.write("Selected") %>>幼圆</option>
            <option value="Andale Mono" <% if rs("SY_FontName")="Andale Mono" Then Response.write("Selected") %>>Andale Mono</OPTION> 
            <option value="Arial"  <% if rs("SY_FontName")="Arial" Then Response.write("Selected") %>>Arial</OPTION> 
            <option value="Arial Black"  <% if rs("SY_FontName")="Arial Black" Then Response.write("Selected") %>>Arial Black</OPTION> 
            <option value="Book Antiqua"  <% if rs("SY_FontName")="Book Antiqua" Then Response.write("Selected") %>>Book Antiqua</OPTION>

        </SELECT></td>
                </tr>
               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">文字大小：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="SY_FontSize" value="<%=rs("SY_FontSize")%>" size="10" <%=onKeyUp%>>
                  像素</td>
                </tr>
               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">文字颜色：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><INPUT name="SY_FontColor" type="text" style="background:<%=rs("SY_FontColor")%>" onClick="Getcolor(this)" value="<%=rs("SY_FontColor")%>" size="7" maxlength="7"></td>
                </tr>
               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">是否粗体：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><SELECT name='SY_Bold' >
				  <OPTION value='1' <% if rs("SY_Bold")=1 Then Response.write("Selected") %>>是</OPTION>
                  <OPTION value='0' <% if rs("SY_Bold")=0 Then Response.write("Selected") %>>否</OPTION>
            
          </SELECT></td>
                </tr>


               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">图片水印文件名：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input name='SY_FileName' type='text' value='<%=rs("SY_FileName")%>' size='35' maxlength='40'> 里请填写图片文件的相对路径，以“\”开头</td>
                </tr>
               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">图片水印宽度：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="SY_Width" value="<%=rs("SY_Width")%>" size="10" <%=onKeyUp%>>
                  像素（会根据宽度自动调整高度,建议150*60透明）</td>
                </tr>
 
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">图片水印透明度：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="SY_Transparence" value="<%=rs("SY_Transparence")%>" size="10" <%=onKeyUp%>>
                  0－1范围，1表示不透明</td>
                </tr>
				  <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">图片水印底色：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><INPUT name="SY_BackgroundColor" type="text" style="background:<%=rs("SY_BackgroundColor")%>" onClick="Getcolor(this)" value="<%=rs("SY_BackgroundColor")%>" size="7" maxlength="7">
                  若想去除水印图片的底色，请在此填入底色的RGB值</td>
                </tr>
               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">图片水印坐标起点：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><select name="SY_coordinates" size="1">
	  <option value="0" <%if rs("SY_coordinates")=0 then%>selected="selected"<%end if%>>左上</option>
	  <option value="1" <%if rs("SY_coordinates")=1 then%>selected="selected"<%end if%>>左下</option>
	  <option value="2" <%if rs("SY_coordinates")=2 then%>selected="selected"<%end if%>>居中</option>
	  <option value="3" <%if rs("SY_coordinates")=3 then%>selected="selected"<%end if%>>右上</option>
	  <option value="4" <%if rs("SY_coordinates")=4 then%>selected="selected"<%end if%>>右下</option>
	  </select></td>
                </tr>
                <tr bgcolor="#BDBEDE">
                  <td height="28" colspan="2" style="border-top-style: none; border-top-width: medium; border-bottom-style: solid; border-bottom-width: 1">
                    <span class="style1"><font size="2">&nbsp;在线客服设定</font></span></td>
                </tr>

               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">客服类型：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><select name="kefu" style="BACKGROUND-COLOR: #FFCCFF; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;">
                            <option value="0" <% if rs("kefu")=0 Then Response.write("Selected") %>>关闭客服</option>                            
                            <option value="1" <% if rs("kefu")=1 Then Response.write("Selected") %>>腾讯QQ</option>
							<option value="2" <% if rs("kefu")=2 Then Response.write("Selected") %>>53快服</option>
							<option value="3" <% if rs("kefu")=3 Then Response.write("Selected") %>>54客服</option>
                            <option value="4" <% if rs("kefu")=4 Then Response.write("Selected") %>>53快服+腾讯QQ</option>
                          </select>
                  </td>
                </tr>

               <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">QQ号码：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="S10" value="<%=Application(CacheName&"_WebSetting")(10)%>" size="30">
                  多个QQ号用西文逗号隔开，QQ号与下面的昵称一一对应</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">QQ对应昵称：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="S11" value="<%=Application(CacheName&"_WebSetting")(11)%>" size="30">
                  多个昵称用西文逗号隔开，昵称与上面的QQ号一一对应</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">QQ显示项目：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="qqshow" value="<%=rs("qqshow")%>" size="30" <%=onKeyUp%>>
                  0QQ号，1昵称，3图标，4QQ号+昵称，5昵称+图标，6QQ号+昵称+图标</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">QQ图标风格：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="qqimg" value="<%=rs("qqimg")%>" size="30" <%=onKeyUp%>>
                  1－13种，<a target=blank href='http://www.tool.la/QQCode/'><font color="#0000FF">样式查看</font></a></td>
                </tr>
                <!--tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">QQ显示位置：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20">
				  <%if rs("qqskin")=1 then%>               
                <input type="radio" name="qqskin" value="1" id="qqskin" checked>
                 左侧&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="qqskin" value="0" id="qqskin">
                右侧
                <%else%>                         
                <input type="radio" name="qqskin" value="1" id="qqskin">
                 左侧&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="qqskin" value="0" id="qqskin" checked>
                右侧<%end if%>&nbsp; <img align=top src=../images/memo.gif alt="客服QQ在屏幕的位置">
                 </td>
                </tr-->

                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">QQ显示框样式：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="radio" name="kefuskin" value="1" <%if rs("kefuskin")=1 then%>checked<%end if%>>样式一	<input type="radio" name="kefuskin" value="2" <%if rs("kefuskin")=2 then%>checked<%end if%>>样式二	<input type="radio" name="kefuskin" value="3" <%if rs("kefuskin")=3 then%>checked<%end if%>>样式三	<input type="radio" name="kefuskin" value="4" <%if rs("kefuskin")=4 then%>checked<%end if%>>样式四	<input type="radio" name="kefuskin" value="5" <%if rs("kefuskin")=5 then%>checked<%end if%>>样式五<a class="info" href=../images/qq/qq.gif target=_blank><img src="../images/memo.gif"  BORDER="0" /><span>点击预览效果</span></a></td>
                </tr>              
                 <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">53快服帐号：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="kefuid" value="<%=rs("kefuid")%>" size="30">
                  <br>到<a target=blank href='http://www.53kf.com'><font color="#0000FF">53快服官网</font></a>申请一个帐号并下载安装客服端即可（需安装客户端，界面友好）</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">54客服代码：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><textarea NAME="k54kefu" ROWS=3 style="overflow:auto;width=70%"><%=rs("k54kefu")%></textarea>
                  <br>到<a target=blank href='http://wwww.54kefu.net/'><font color="#0000FF">54客服官网</font></a>申请一个帐号并获取代码填上即可（可集成多种客服工具，自由发挥）</td>
                 </tr>


                <tr bgcolor="#BDBEDE">
                  <td height="28" colspan="2" style="border-top-style: none; border-top-width: medium; border-bottom-style: solid; border-bottom-width: 1">
                    <span class="style1"><font size="2">&nbsp;虚拟货币与道具设置</font></span></td>
                </tr>
                <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">虚拟货币名称：</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="rmb_mc" size="20" value="<%=rs("rmb_mc")%>">
                    具有网站特点的虚拟货币个性名称，如虚拟币、铜版、玉币</td>
                </tr> 
                <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">
                  <font color="#0000FF">1元人民币</font>可以<font color="#0000FF">购买</font>：</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="rmb_hb" size="20" value="<%=rs("rmb_hb")%>" DISABLED <%=onKeyUp%>>
                    个<%=rs("rmb_mc")%>（为便于虚拟币的购买转换，这个值不要修改了）</td>
                </tr> 

                <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">购买<font color="#990000">标题颜色道具</font>需：</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="g_a" size="20" value="<%=rs("g_a")%>" <%=onKeyUp%>>
                    个<%=rs("rmb_mc")%></td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">购买<font color="#990000">信息置顶道具</font>需：</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="g_b" size="20" value="<%=rs("g_b")%>" <%=onKeyUp%>>
                    个<%=rs("rmb_mc")%></td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">购买<font color="#990000">内容贴图道具</font>需：</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="g_c" size="20" value="<%=rs("g_c")%>" <%=onKeyUp%>>
                    个<%=rs("rmb_mc")%></td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">购买<font color="#990000">通过审核道具</font>需：</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="g_d" size="20" value="<%=rs("g_d")%>" <%=onKeyUp%>>
                    个<%=rs("rmb_mc")%></td>
                </tr>
                <tr bgcolor="#BDBEDE">
                  <td height="28" colspan="2" style="border-top-style: none; border-top-width: medium; border-bottom-style: solid; border-bottom-width: 1">
                    <span class="style1"><font size="2">&nbsp;积分奖励与转换设置</font></span></td>
                </tr>
                <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">注册赠送：</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="jf_1" size="20" value="<%=rs("jf_1")%>" <%=onKeyUp%>>
                    点积分</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">发布信息每条赠送：</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="jf_2" size="20" value="<%=rs("jf_2")%>" <%=onKeyUp%>>
                    点积分（自己删除扣同样分，管理员删除加倍扣分）</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">新闻资讯投稿每条赠送：</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="tgjf" size="20" value="<%=rs("tgjf")%>" <%=onKeyUp%>>
                    点积分</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">登陆一次获取：</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="jf_3" size="20" value="<%=rs("jf_3")%>" <%=onKeyUp%>>
                    点积分 (每天一次)</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">消费<%=rs("rmb_mc")%>可增加：</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="jf_4" size="20" value="<%=rs("jf_4")%>" <%=onKeyUp%>>
                    点积分</td>
                </tr>

                <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">积分转换<%=rs("rmb_mc")%>：</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="jf_hb" size="20" value="<%=rs("jf_hb")%>" <%=onKeyUp%>>
                     点积分可转换1个<%=rs("rmb_mc")%></td>
                </tr> 
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">1个<%=rs("rmb_mc")%>转换：</td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="adjfs" size="20" value="<%=rs("adjfs")%>" <%=onKeyUp%>>
                    点积分</td>
                </tr>
                <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">奖励注册推荐人：</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="jf_5" size="20" value="<%=rs("jf_5")%>" <%=onKeyUp%>>
                    点积分</td>
                </tr>
                <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">查看权限信息联系手机扣：</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="jf_ck" size="20" value="<%=rs("jf_ck")%>" <%=onKeyUp%>>
                    点积分</td>
                </tr>                <tr bgcolor="#BDBEDE">
                  <td height="28" colspan="2" bgcolor="#BDBEDE" style="border-top-style: none; border-top-width: medium; border-bottom-style: solid; border-bottom-width: 1">
                    <span class="style1"><font size="2">&nbsp;会员注册选项</font></span></td>
                </tr>
                <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">赠送的<font color="#0000FF"><%=rs("rmb_mc")%></font><font color="#800000">：</font></td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="z_hb" size="20" value="<%=rs("z_hb")%>" <%=onKeyUp%>>
                    个</td>
                </tr>
                <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">赠送的<font color="#800000">标题变色道具：</font></td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="z_a" size="20" value="<%=rs("z_a")%>" <%=onKeyUp%>>
                    个</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">赠送的<font color="#800000">信息置顶道具：</font></td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="z_b" size="20" value="<%=rs("z_b")%>" <%=onKeyUp%>>
                    个</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">赠送的<font color="#800000">内容贴图道具：</font></td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="z_c" size="20" value="<%=rs("z_c")%>" <%=onKeyUp%>>
                    个</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">赠送的<font color="#800000">通过审核道具：</font></td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="z_d" size="20" value="<%=rs("z_d")%>" <%=onKeyUp%>>
                    个 （慎重给审核道具）</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">注册同时发送邮件通知：</font></td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="regmail" size="20" value="<%=rs("regmail")%>" <%=onKeyUp%>> （0=不发送，1=发送）</td>
                </tr>
				                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">注册会员是否要邮件确认：</font></td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="mailqr" size="20" value="<%=rs("mailqr")%>" <%=onKeyUp%>> （0=不需要，1=需要）</td>
                </tr>
				  <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">注册会员是否要审核：</font></td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="usersh" size="20" value="<%=rs("usersh")%>" <%=onKeyUp%>> （0=不审核，1=审核）</td>
                </tr>
				  <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">邮件确认的同时自动审核：</font></td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="zdsh" size="20" value="<%=rs("zdsh")%>" <%=onKeyUp%>> （0=不自动审核，1=自动审核）</td>
                </tr>
				  <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">邮件确认有效时间：</font></td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="regyxq" size="20" value="<%=rs("regyxq")%>" <%=onKeyUp%>>天 （超过天数即自动删除临时注册资料）</td>
                </tr>
				 <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">手机短信验证开关：
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"> <%if rs("sms_kg")=True then%>               
                <input type="radio" name="sms_kg" value="1" id="sms_kg" checked>
                 启用&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="sms_kg" value="0" id="sms_kg">
                关闭
                <%else%>                         
                <input type="radio" name="sms_kg" value="1" id="sms_kg">
                 启用&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="sms_kg" value="0" id="sms_kg" checked>
                关闭<%end if%> （如果没有帐号要关闭，否则会员无法注册）</td>
                </tr>
				  <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">手机短信验证帐号：</font></td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="sms_user" size="20" value="<%=rs("sms_user")%>"> （到<a target=blank href='http://www.dxton.com/'><font color="#0000FF">短信通</font></a>申请一个帐号并填写有关参数。<a target=blank href='http://www.dxton.com/webservice/sms.asmx/GetNum?account=<%=rs("sms_user")%>&password=<%=rs("sms_pass")%>'><font color="#0000FF">查询帐户余额</font></a>）</td>
                </tr>
				  <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">手机短信验证密码：</font></td>
                  <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="password" name="sms_pass" size="20" value="<%=rs("sms_pass")%>"> （此密码要与短信通平台的接口密码一致）</td>
                </tr> 
                 <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">手机短信验证内容：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20">您的验证码是:****。<textarea name="sms_content" cols="55" rows="2" class="input2"><%=rs("sms_content")%></textarea> <br>（<span class="style2">重要：如果申请的手机验证帐号不是VIP3以上的，此内容绝对不能增减或修改任何字符，否则无法通过短信通平台审核而无法发送短信！</span>）</td>
                </tr>
				 <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">手机号能否重复注册：
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"> <%if rs("sms_cs")=True then%>               
                <input type="radio" name="sms_cs" value="1" id="sms_cs" checked>
                 不能&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="sms_cs" value="0" id="sms_cs">
                可以
                <%else%>                         
                <input type="radio" name="sms_cs" value="1" id="sms_cs">
                 不能&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="sms_cs" value="0" id="sms_cs" checked>
                可以<%end if%> （建议不准重复注册）</td>
                </tr>
                 <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">同一IP注册验证次数：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="sms_ip" size="5" value="<%=rs("sms_ip")%>" <%=onKeyUp%>> (0为不限制)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;取消同IP验证次数限制时间 <input type="text" name="sms_time" size="4" value="<%=rs("sms_time")%>" <%=onKeyUp%>> 小时后(0为不限制)</td>
                </tr>
                 <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">自动删除验证记录时间：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="sms_del" size="5" value="<%=rs("sms_del")%>" <%=onKeyUp%>> 天 （0为不删除。建议设置超过30天自动删除。删除后将不能限制已注册过的手机和IP地址，根据需要设置。）</td>
                </tr>
                <tr bgcolor="#BDBEDE">
                  <td height="28" colspan="2" style="border-top-style: none; border-top-width: medium; border-bottom-style: solid; border-bottom-width: 1">
                    <span class="style1"><font size="2">&nbsp;VIP设定</font></span></td>
                </tr>
                 <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">VIP资格有效期至少申请：</span></td>
                  <td style="border-bottom-style: none; border-bottom-width: medium" height="29" bgcolor="#FFFFFF"><p style="margin-left: 20" >
                      <input type="text" name="vipsj" value="<%=rs("vipsj")%>" size="20"<%=onKeyUp%> />
个月 (建议时间长点，便于管理)</td> 
                </tr>               
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">VIP会员资格需要：</span></td>
                  <td style="border-bottom-style: none; border-bottom-width: medium" height="29" bgcolor="#FFFFFF"><p style="margin-left: 20">
                      <input type="text" name="vipje" value="<%=rs("vipje")%>" size="20"  <%=onKeyUp%>/>
元人民币/月 (必须为大于0的整数)</td>
                </tr>

                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">非VIP用户只显示：</span></td>
                  <td style="border-bottom-style: none; border-bottom-width: medium" height="29" bgcolor="#FFFFFF"><p style="margin-left: 20">
            <input type="text" name="huiyuansp" value="<%=rs("huiyuansp")%>" size="20"  <%=onKeyUp%>/>
个店铺商品</td>
                </tr>
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">非VIP用户只显示：</span></td>
                  
				  
				  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="huiyuanxx" type="text" id="huiyuanxx" value="<%=rs("huiyuanxx")%>" size="20" <%=onKeyUp%> />
				    条个人信息</p></td>
                </tr>
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">申请店企至少要发布：</span></td>
				  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="xinxis" type="text" id="xinxis" value="<%=rs("xinxis")%>" size="20" <%=onKeyUp%> />
				    条供求信息</p></td>
                </tr>
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">VIP会员店铺文章发布限：</span></td>
				  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="usernews" type="text" id="usernews" value="<%=rs("usernews")%>" size="20" <%=onKeyUp%> />
				    篇（0不限制）</p></td>
                 </tr>
                 <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">VIP会员店铺友情链接限：</span></td>
				  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="userlink" type="text" id="userlink" value="<%=rs("userlink")%>" size="20" <%=onKeyUp%> />
				    个（0不限制）</p></td>
                </tr>
                 <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">VIP店铺屏蔽广告积分下限：</span></td>
				  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="adjf" type="text" id="adjf" value="<%=rs("adjf")%>" size="20" <%=onKeyUp%> />
				    分（少于则显示网站广告，0始终显示）</p></td>
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
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">申请店铺名片要审核的证照：</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20"><INPUT TYPE=checkbox NAME="certificate1" VALUE="1" <%if certificate1="1" then response.write("checked")%>>法人身份证&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="certificate2" VALUE="1" <%if certificate2="1" then response.write("checked")%>>法人标准像&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="certificate3" VALUE="1" <%if certificate3="1" then response.write("checked")%>>营业执照&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="certificate4" VALUE="1" <%if certificate4="1" then response.write("checked")%>>国税登记证&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="certificate5" VALUE="1" <%if certificate5="1" then response.write("checked")%>>地税登记证<br><INPUT TYPE=checkbox NAME="certificate6" VALUE="1" <%if certificate6="1" then response.write("checked")%>>组织机构代码证
					&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="certificate7" VALUE="1" <%if certificate7="1" then response.write("checked")%>>其它证照(门面、车间、商标等)
&nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>会员申请店企时可看到的审核条件</span></a></td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">申请店企条件提醒：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><textarea name="prompt" cols="70" rows="3" class="input2" onpropertychange="kn()"><%=rs("prompt")%></textarea>
                  &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>提醒会员申请店企应注意的条件。限200字符内。</span></a><br><span id=num>您还可以输入200个字节，100汉字</span></td>
                </tr>
                <tr bgcolor="#BDBEDE">
                  <td height="28" colspan="2" bgcolor="#BDBEDE" style="border-top-style: none; border-top-width: medium; border-bottom-style: solid; border-bottom-width: 1"><span class="style1"><font size="2">&nbsp;发布信息方面设置</font></span></td>
                </tr>

                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">是否注册才可发布：</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="addxinxi" type="text" value="<%=rs("addxinxi")%>" size="20" maxlength="1" <%=onKeyUp%> />
                    0=注册会员才能发布，1=无需注册直接发布</p></td>
                </tr>
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">注册后多久才能发布：</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="zcfbsj" type="text" value="<%=rs("zcfbsj")%>" size="20" maxlength="2" <%=onKeyUp%> />
                    秒&nbsp;&nbsp;为防止群发垃圾信息，可设60秒以上</p></td>
                </tr>
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right"><span class="style2" style="margin-left: 20">游客发布信息设置：</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="xinxish" type="text" value="<%=rs("xinxish")%>" size="20" maxlength="1" <%=onKeyUp%> />
                    0=等待审核才能通过，1=无需审核直接通过 <font color="#ff0000">谨慎使用</font></p></td>
                </tr>
                 <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right"><span class="style2" style="margin-left: 20">非VIP发布信息设置：</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="vipno" type="text" value="<%=rs("vipno")%>" size="20" maxlength="1" <%=onKeyUp%> />
                    0=等待审核才能通过，1=无需审核直接通过 <font color="#ff0000">谨慎使用</font></p></td>
                </tr>
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right"><span class="style2" style="margin-left: 20">发布信息审核总开关：</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="onoff" type="text" value="<%=rs("onoff")%>" size="20" maxlength="1" <%=onKeyUp%> />
                    0=等待审核才能通过，1=按上方设置条件通过 <font color="#ff0000">谨慎使用</font></p></td>
                </tr>
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">信息标题长度限制：</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="biaotinum" type="text" value="<%=rs("biaotinum")%>" size="20" <%=onKeyUp%>  />
                    个字节（一个汉字2个字节），建议40字节内</p></td>
                </tr>
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">信息说明长度限制：</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="memonum" type="text" value="<%=rs("memonum")%>" size="20" <%=onKeyUp%>  />
                    个字节（一个汉字2个字节），建议2000字节内</p></td>
              </tr>
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">更新信息扣除积分：</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="gxkf" type="text" value="<%=rs("gxkf")%>" size="20" <%=onKeyUp%>  />
                    分 会员一天后更新信息（更新后排行靠前）将扣除的积分，0为不扣除</p></td>
                </tr>			
                <tr bgcolor="#BDBEDE">
                  <td height="28" colspan="2" style="border-top-style: none; border-top-width: medium; border-bottom-style: solid; border-bottom-width: 1">
                    <span class="style1"><font size="2">&nbsp;信息显示设置</font></span></td>
                </tr>
                <!--<tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">是否生成静态首页：</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="HtmlHome" type="text" value="<%=rs("HtmlHome")%>" size="20" maxlength="1" <%=onKeyUp%> />
                      <label></label>
                    0=asp动态首页方式显示，1=HTM静态首页方式显示<br><font color="#ff0000">(如果选择按地区（分站）方式显示有关内容请勿选择静态首页，原因是不静态不能自动按地区显示信息。)</font></p></td>
                </tr>-->
				<input type="hidden" name="HtmlHome" value="0">
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">是否生成静态内容页：</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="asphtm" type="text" value="<%=rs("asphtm")%>" size="20" maxlength="1" <%=onKeyUp%> />
                      <label></label>
                    0=asp动态页方式显示，1=HTM静态页方式显示<br><font color="#ff0000">(静态内容页有利访问速度、搜索引擎收录和减轻服务器压力，但占用空间，用户视条件决定)</font></p></td>
                </tr>
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">过滤前台发布信息内容HTM代码：</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="Filterhtm" type="text" value="<%=rs("Filterhtm")%>" size="20" <%=onKeyUp%>  />
                    0不过滤，1过滤游客，2过滤非VIP会员，3全部过滤</p></td>
                </tr>
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">信息详细内容显示长度：</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="content_length" type="text" value="<%=rs("content_length")%>" size="20" <%=onKeyUp%> />
                      <label></label>
                    个字节  0为不显示内容，其它按设定字节长度，一个汉字为2个字节<br><font color="#ff0000">(当发布信息选择阅读权限时起作用)</font></p></td>
                </tr>
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">完成交易是否显示：</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="jiaoyi" type="text" value="<%=rs("jiaoyi")%>" size="20" maxlength="1" <%=onKeyUp%> />
                      <label></label>
                    0=不显示已交易完成，1=同时显示交易完成</p></td>
                </tr>
                <tr>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium" align="right">过期信息是否显示：</span></td>
                  <td height="29" bgcolor="#FFFFFF" style="border-bottom-style: none; border-bottom-width: medium"><p style="margin-left: 20">
                      <input name="overdue" type="text" value="<%=rs("overdue")%>" size="20" maxlength="1" <%=onKeyUp%> />
                      <label></label>
                    0=不显示过期信息，1=同时显示过期信息</p></td>
                </tr>

                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">使用置顶道具有效期：</td>
                 <td width="70%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="b_y" size="20" value="<%=rs("b_y")%>" <%=onKeyUp%>> 
天&nbsp;&nbsp;系统将取消符合条件的信息置顶(0为不受限制，<font color="#ff0000">谨慎使用</font>)</td>
                </tr>
                  <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">推荐时间有效期：</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="tui_y" size="20" value="<%=rs("tui_y")%>" <%=onKeyUp%>> 
天&nbsp;&nbsp;系统将取消符合条件的信息推荐(0为不受限制，<font color="#ff0000">谨慎使用</font>)</td>
                </tr>
                  <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">热门信息标准：</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="hotbz" size="20" value="<%=rs("hotbz")%>" <%=onKeyUp%>> 
次&nbsp;&nbsp;热门信息需要达到的点击数</td>
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
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">首页头条图片标题参数：</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20"><INPUT type="text" NAME="HomeElite1" size="5" VALUE="<%=HomeElite1%>" <%=onKeyUp%>>图片宽度 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <INPUT type="text" NAME="HomeElite2" size="5" VALUE="<%=HomeElite2%>" <%=onKeyUp%>>图片高度 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT type="text" NAME="HomeElite3" size="5" VALUE="<%=HomeElite3%>" <%=onKeyUp%>>字体大小 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <br><INPUT type="text" NAME="HomeElite4" size="15" VALUE="<%=HomeElite4%>">字 体<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>服务器上必须有此字体，否则生成时采用默认的宋体！</span></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT type="text" NAME="HomeElite5" size="15" VALUE="<%=HomeElite5%>">存储位置<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>图片存储位置例如UploadFiles/Home/，不可以“/”或者“..”开头。目录要事先手工建立！</span></a>&nbsp;&nbsp;<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>如首页相应位置规格不改变请勿修改头条图片标题参数！</span></a></td>
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
                    <span class="style1"><font size="2">&nbsp;验证方案设置</font></span></td>
                </tr>
				<tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">选择验证方案：</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20"><INPUT TYPE=checkbox NAME="codekg1" VALUE="1" <%if codekg1="1" then response.write("checked")%>>问答验证&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="codekg2" VALUE="1" <%if codekg2="1" then response.write("checked")%>>数字验证&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="codekg3" VALUE="1" <%if codekg3="1" then response.write("checked")%>>文字验证&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="codekg4" VALUE="1" <%if codekg4="1" then response.write("checked")%>>星期验证
					&nbsp;&nbsp;<INPUT TYPE=checkbox NAME="codekg5" VALUE="1" <%if codekg5="1" then response.write("checked")%>>算式验证
&nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>有四种方式验证，可多选和经常变换防灌水。建议至少选用问答验证码！</span></a></td>
                </tr>
                <tr bgcolor="#BDBEDE">
                  <td height="28" colspan="2" style="border-top-style: none; border-top-width: medium; border-bottom-style: solid; border-bottom-width: 1">
                    <span class="style1"><font size="2">&nbsp;其它设置</font></span></td>
                </tr>
                  <tr>
                  <td width="30%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF" align="right">地区分类导航显示个数：</td>
                  <td width="70%" style="border-bottom-style: none; border-bottom-width: medium" height="25" bgcolor="#FFFFFF">
                    <p style="margin-left: 20">
                    <input type="text" name="citys" size="20" value="<%=rs("citys")%>" <%=onKeyUp%>> 
个&nbsp;&nbsp;可根据分类名称长度适当设置</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">首页幻灯广告显示类别：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="radio" name="slidelx" value="0" <%if rs("slidelx")=0 then%>checked<%end if%>>幻灯广告A（分地区显示） <input type="radio" name="slidelx" value="1" <%if rs("slidelx")=1 then%>checked<%end if%>>幻灯广告B（仅显示总站广告） </td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">首页视频显示类别：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="radio" name="hdlb" value="1" <%if rs("hdlb")=1 then%>checked<%end if%>>普通视频广告 <input type="radio" name="hdlb" value="2" <%if rs("hdlb")=2 then%>checked<%end if%>>FLV视频广告 <input type="radio" name="hdlb" value="0" <%if rs("hdlb")=0 then%>checked<%end if%>>关闭视频广告(将播放标准广告ID:240)</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">新闻资讯投稿开关：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="radio" name="contribute" value="1" <%if rs("contribute")=1 then%>checked<%end if%>>游客 <input type="radio" name="contribute" value="2" <%if rs("contribute")=2 then%>checked<%end if%>>普通会员 <input type="radio" name="contribute" value="3" <%if rs("contribute")=3 then%>checked<%end if%>>VIP会员 <input type="radio" name="contribute" value="0" <%if rs("contribute")=0 then%>checked<%end if%>>关闭投稿入口 （级别高向下兼容）</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">信息回复、新闻评论开关：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="Newspl" value="<%=rs("Newspl")%>" size="20">
                  &nbsp;&nbsp; 0为全开放，1为需要登录才能评论，2为显示评论禁止新评论，3为全部禁止</td>
                </tr>
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">信息回复、新闻评论审核：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="plsh" value="<%=rs("plsh")%>" size="20">
                  &nbsp;&nbsp; 0为直接显示，1为审核显示</td>
                </tr>				
                <tr>
                  <td width="30%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF" align="right">网站缓存名称：</td>
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20"><input type="text" name="CacheName" value="<%=rs("CacheName")%>" size="20">
                  &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>随意填一个不带空格的英文名称即可，不能使用汉字！可以使用系统默认的。</span></a></td>
                </tr>
			
              </table>
            </td>
          </tr>
 

          <tr>
            <td align="center" width="688" bgcolor="#eeeeee" height="35">            <table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><div align="center">
                  <input type="submit" value="确认设置" name="B1">
                </div></td>
                <td><div align="center">
                  <input type="reset" value="取消设置" name="B1">
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
