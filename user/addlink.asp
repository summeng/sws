<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file="../Include/md5.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>

<%Call OpenConn
dim action,url,wzname,name,mail,info,img,dd,adminlink,ttt
id=strint(request("id"))
If request.form("id")="" Then id=strint(request("id"))
action=HtmlEncode(request.querystring("action"))
name=trim(request("name"))
mail=trim(request("mail"))
info=trim(request("info"))
wzname=HtmlEncode(request("wzname"))
url=request("url")
adminlink=trim(request("adminlink"))
city_oneid=strint(request.form("city_one"))
city_twoid=strint(request.form("city_two"))
city_threeid=strint(request.form("city_three"))
if city_oneid>0 then
set rs=server.createobject("adodb.recordset")
rs.open "select city from DNJY_city where twoid=0 and threeid=0 and id="&city_oneid,conn,1,1
city_one=rs("city")
rs.close
if city_twoid>0 then
rs.open "select city from DNJY_city where id="&city_oneid&" and threeid=0 and twoid="&city_twoid,conn,1,1
city_two=rs("city")
rs.close
end if
if city_twoid>0 and city_threeid>0 then
rs.open "select city from DNJY_city where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid,conn,1,1
city_three=rs("city")
rs.close
end if
set rs=nothing
end If

select case action
case "add"
set rs=server.CreateObject("adodb.recordset")
	rs.open "select wzname from DNJY_link where name="""&name&"""",conn,1,1
	if not rs.eof Then
	Response.Write ("<script language=javascript>alert('用户名"&name&"已经存在，请重新选择一个用户名!');history.go(-1);</script>")
	rs.close
	set rs=Nothing
    response.end	
	end If	
		
set rs=server.CreateObject("adodb.recordset")
rs.open "select wzname,url,logo,img,addtime,name,pass,city_oneid,city_twoid,city_threeid,city_one,city_two,city_three,mail,info from DNJY_link",conn,1,3
rs.AddNew
rs(0)=wzname
rs(1)=url
if city_oneid="" then city_oneid=0
if city_twoid="" then city_twoid=0
if city_threeid="" then city_threeid=0
rs("city_oneid")=city_oneid
rs("city_one")=city_one
rs("city_two")=city_two
rs("city_twoid")=city_twoid
rs("city_three")=city_three
rs("city_threeid")=city_threeid
rs("logo")=trim(request("logo"))
If request("logo")="" Then
rs("logo")=0
Else
rs("logo")=trim(request("logo"))
End if
rs("name")=trim(request("name"))
rs("mail")=trim(request("mail"))
rs("info")=trim(request("info"))
rs("pass")=md5(trim(request("pass")))
rs("addtime")=now()

rs.Update
response.Write "<script language=javascript>alert('添加成功，审核后显示!');location.replace('addlink.asp?"&C_ID&"');</script>"
rs.Close:set rs=nothing

case "edit"
If trim(request("logo"))="" And trim(request("img"))=1 Then response.Write "<script language=javascript>alert('选择图片方式须填写图片链接或上传LOGO！');location.replace('?adminlink=1&id="&id&"');</script>"
set rs=server.CreateObject("adodb.recordset")
rs.open "select wzname,url,logo,img,addtime,city_oneid,city_twoid,city_threeid,city_one,city_two,city_three,pass,name,mail,info from DNJY_link where id="&id,conn,1,3
If rs("name")<>trim(request("name")) Then
response.Write "<script language=javascript>alert('您的管理用户名不对!');location.replace('?adminlink=1&id="&id&"');</script>"
ElseIf rs("pass")<>md5(trim(request("pass"))) then
response.Write "<script language=javascript>alert('您的管理密码不对!');location.replace('?adminlink=1&id="&id&"');</script>"
else
rs(0)=wzname
rs(1)=url
if city_oneid="" then city_oneid=0
if city_twoid="" then city_twoid=0
if city_threeid="" then city_threeid=0
rs("city_oneid")=city_oneid
rs("city_one")=city_one
rs("city_two")=city_two
rs("city_twoid")=city_twoid
rs("city_three")=city_three
rs("city_threeid")=city_threeid
If request("logo")="" Then
rs("logo")=0
Else
rs("logo")=trim(request("logo"))
End if
rs("img")=trim(request("img"))
rs("mail")=trim(request("mail"))
rs("info")=trim(request("info"))
rs.Update
rs.Close:set rs=Nothing
response.Write "<script language=javascript>alert('修改成功!');location.replace('?adminlink=1&id="&id&"');</script>"
End if

case "del"
Dim Num,objFSO
if action="del" then
If Request.Form("DelID")="" or isnull(Request.Form("DelID")) or isempty(Request.Form("DelID")) then
response.write "<font style='color:red;font-size:12px'>请选择要删除的信息</font>"
else
Num=request.form("DelID").count
for i=1 to Num
set rs = server.CreateObject ("Adodb.recordset")
rs.open "select * from DNJY_link where id="&request.form("DelID")(i),conn,3,3
do while not rs.eof

logo=trim(rs("logo"))
if left(logo,6)="images" then
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if objFSO.fileExists(Server.MapPath("../"&logo)) then
objFSO.DeleteFile(Server.MapPath("../"&logo))
end if
set objfso=nothing
end If

rs.delete
rs.movenext
Loop
Next
Call CloseRs(rs)
end If
response.write "<script language='javascript'>" & chr(13)
		response.write "alert('成功删除！');" & chr(13)
		response.write "window.document.location.href='addlink.asp?"&C_ID&"';"&chr(13)
		response.write "</script>" & chr(13)
response.end
end If

end select
%>
<script LANGUAGE=JavaScript>
  function textLimitCheck(thisArea, maxLength){
    if (thisArea.value.length > maxLength){
      alert(maxLength + ' 个字限制. \r超出的将自动去除.');
      thisArea.value = thisArea.value.substring(0, maxLength);
      thisArea.focus();
    }
    /*回写span的值，当前填写文字的数量*/
    messageCount.innerText = thisArea.value.length;
  }
</script>
<script language = "JavaScript">
<!--
//字节数限制，与下面配合=========
String.prototype.Length=function(){
temp=this.replace(/([^\x00-\xff])/g,"$1$1");
return temp.length
}
//显示输入字符数

function kn()
  {
   var ln=form.wzname.value.Length();
     num.innerText='您还可以输入:'+(16-ln)+'个字符';
      //if (ln>=16) form.wzname.readOnly=true;  //这行是如果满足条件表单无法输入和修改
}
function kn2()
 {
   var ln=form.name.value.Length();
     num2.innerText='您还可以输入:'+(8-ln)+'个字符';
      
}
function kn3()
 {
   var ln=form.pass.value.Length();
     num3.innerText='您还可以输入:'+(8-ln)+'个字符';
      
}
//验证网址正确性
function checkeURL(URL){ 
var str=URL;
//在JavaScript中，正则表达式只能使用"/"开头和结束，不能使用双引号
//判断URL地址的正则表达式为:http(s)?://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?
//下面的代码中应用了转义字符"\"输出一个字符"/"
var Expression=/http(s)?:\/\/([\w-]+\.)+[\w-]+(\/[\w- .\/?%&=]*)?/; 
var objExp=new RegExp(Expression);
if(objExp.test(str)==true){
return true;
}else{
return false;
}
}

//验证邮箱正确性
function checkemail(mail){
var str=mail;
//在JavaScript中，正则表达式只能使用"/"开头和结束，不能使用双引号
var Expression=/\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*/; 
var objExp=new RegExp(Expression);
if(objExp.test(str)==true){
return true;
}else{
return false;
}
} 

function CheckForm(){
   
	if (document.form.wzname.value.length == 0) {
		alert("请填网站名称！");
		document.form.wzname.focus();
		return false;
	}
     if ((document.form.wzname.value.Length()>16) || (document.form.wzname.value.Length()<1)) //字节数限制，与上面配合
     {
	 alert("网站名称长度要求1－16字节，请重新输入！");
	  document.form.wzname.focus()
	  return false
  }
	if (document.form.url.value.length == 0) {
		alert("请填网址！");
		document.form.url.focus();
		return false;
	}
    var url=form.url.value;
     if(!checkeURL(url)){
     alert("您输入的网址不符合规范，请重新输入！");
	  document.form.url.focus()
	  return false
     }
	if (document.form.name.value.length == 0) {
		alert("请填管理用户名！");
		document.form.name.focus();
		return false;
	}
  if ((document.form.name.value.Length()>8) || (document.form.name.value.Length()<1)) //字节数限制，与上面配合
     {
	 alert("用户名长度要求1－8字节，请重新输入！");
	  document.form.name.focus()
	  return false
  }
	if (document.form.pass.value.length == 0) {
		alert("请填管理密码！");
		document.form.pass.focus();
		return false;
	}
<%if adminlink="2" then%>
if (document.form.passs.value=="")
	{
      alert("确认密码不能为空，请重新输入！");
	  document.form.passs.focus()
	  return false
	 }
    if(document.form.pass.value != document.form.passs.value) {
	document.form.passs.focus();
	document.form.passs.value = '';
    alert("两次输入的密码不同，请重新输入！");
	return false;
  }
<%end if%>
	if (document.form.mail.value.length == 0) {
		alert("请填正确的电子信箱！");
		document.form.mail.focus();
		return false;
	}
    var mail=form.mail.value;
     if(!checkemail(mail)){
     alert("您输入的信箱不符合规范，请重新输入！");
	  document.form.mail.focus()
	  return false
     }

}
// -->
</script>
<HTML>
<HEAD>
<TITLE>友情链接管理</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK rel="stylesheet" href="/<%=strInstallDir%>css/style.css" type="text/css">
</HEAD>
<BODY>
<TABLE width="600" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">

 <%If adminlink="1" Then
  set rs=server.CreateObject("adodb.recordset")
  rs.Open "select * from DNJY_link where id="&id,conn,1,1%>
		  <FORM name="form" method="post" action="?action=edit&id=<%=int(rs("id"))%>&city_oneid=<%=city_oneid%>&city_twoid=<%=city_twoid%>&city_threeid=<%=city_threeid%>" language="javascript" onsubmit="return CheckForm()">
		   <TR> 
    <TD height="20" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">修 改 友 情 链 接</FONT></TD>
  </TR>
  <TR> 
    <TD bgcolor="#FFFFFF"> <BR>
	<TABLE width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
            <tr> 
      <td width="100" height="25">所属地区：</td>
      <td colspan="2">
<%Dim rsi
set rsi=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")
%>
      <SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
		dim count:count = 0
        do while not rsi.eof 
        %>
subcat[<%=count%>] = new Array("<%=rsi("city")%>","<%=rsi("id")%>","<%=rsi("twoid")%>");
        <%
        count = count + 1
        rsi.movenext
        loop
        rsi.close
		set rsi=nothing
        %>
onecount=<%=count%>;
</SCRIPT>
<%
set rsi=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid>0")
%>
 <SCRIPT language = "JavaScript">
var onecount4;
onecount4=0;
subcat4 = new Array();
        <%
		dim count4:count4 = 0
        do while not rsi.eof 
        %>
subcat4[<%=count4%>] = new Array("<%=rsi("city")%>","<%=rsi("id")%>","<%=rsi("twoid")%>","<%=rsi("threeid")%>");
        <%
        count4 = count4 + 1
        rsi.movenext
        loop
        rsi.close
		set rsi = nothing
        %>
onecount4=<%=count4%>;

function changelocation(locationid)
    {
    document.form.city_two.length = 0; 
	document.form.city_two.options[0] = new Option('选择城市','');
	document.form.city_three.length = 0; 
	document.form.city_three.options[0] = new Option('选择城市','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.form.city_two.options[document.form.city_two.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
	
	function changelocation4(locationid,locationid1)
    {
    document.form.city_three.length = 0; 
    document.form.city_three.options[0] = new Option('选择城市','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount4; i++)
        {
            if (subcat4[i][2] == locationid)
            { 
			if (subcat4[i][1] == locationid1)
			{
                document.form.city_three.options[document.form.city_three.length] = new Option(subcat4[i][0], subcat4[i][3]);
            }        
        }
       } 
    }            
              </SCRIPT>
      <SELECT name="city_one" size="1" id="select2" onChange="changelocation(document.form.city_one.options[document.form.city_one.selectedIndex].value)">
            <OPTION value="" selected>选择城市</OPTION>
  <%set rsi=conn.execute("select * from DNJY_city where id>0 and twoid=0 and threeid=0 order by indexid asc")
if rsi.eof or rsi.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsi.eof%>
  <OPTION value="<%=rsi("id")%>" <%if rsi("id")=rs("city_oneid") then%>selected<%end if%>><%=rsi("city")%></OPTION>
 <%rsi.movenext
    loop
	%>
	<%end if
rsi.close
set rsi = nothing
%>
      </SELECT> 
	  <SELECT name="city_two" id="city_two" onChange="changelocation4(document.form.city_two.options[document.form.city_two.selectedIndex].value,document.form.city_one.options[document.form.city_one.selectedIndex].value)">
    <OPTION value="" selected>选择城市</OPTION>
   <%
set rsi=conn.execute("select * from DNJY_city where id="&rs("city_oneid")&" and twoid>0 and threeid=0")
if rsi.eof and rsi.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsi.eof%>
  <OPTION value="<%=rsi("twoid")%>" <%if rsi("twoid")=rs("city_twoid") then%>selected<%end if%>><%=rsi("city")%></OPTION>
 <%rsi.movenext
    loop
	%>
	<%end if
rsi.close
set rsi = nothing
%>
	</SELECT>
	     <SELECT name="city_three" id="city_three">
         <OPTION value="" selected>选择城市</OPTION>
		<%
set rsi=conn.execute("select * from DNJY_city where id="&rs("city_oneid")&" and twoid="&rs("city_twoid")&" and threeid<>0")
if rsi.eof and rsi.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsi.eof%>
<OPTION value="<%=rsi("threeid")%>" <%if rsi("threeid")=rs("city_threeid") then%>selected<%end if%>><%=rsi("city")%></OPTION>
   <% rsi.movenext
    loop
	%>
<%end if
rsi.close
set rsi = nothing
%>
    </SELECT></td>
    </tr>
  <input type="hidden" name="ttt" id="ttt" value="1">
  <input type="hidden" name="id" id="id" value="<%=rs(2)%>">
    <tr> 
      <td width="100" height="25"><font color="#FF0000">*</font>网站名称</td>
      <td colspan="2">
<INPUT name="wzname" type="text" id="wzname" value="<%=trim(rs("wzname"))%>" size="20"   maxlength="16" onpropertychange="kn()"> <span id=num>你还可以输入16个字符</span> </td>
    </tr>
    <tr> 
      <td width="100" height="25"><font color="#FF0000">*</font>网址</td>
      <td colspan="2">
<INPUT name="url" type="text" id="url" value="<%=trim(rs("url"))%>" size="40"> 前面要带http://</td>
    </tr>
    <tr> 
      <td width="100" height="25"><font color="#FF0000">*</font>管理用户名</td>
      <td colspan="2"><input type="text" name="name" value="" >
</td>
    </tr>
    <tr> 
      <td width="100" height="25"><font color="#FF0000">*</font>管理密码</td>
      <td colspan="2">
<INPUT name="pass" type="password" id="pass" value="" size="21"></td>
    </tr>
    <tr> 
      <td width="100" height="25"><font color="#FF0000">*</font>电子信箱</td>
      <td colspan="2">
<INPUT name="mail" type="text" id="mail" value="<%=trim(rs("mail"))%>" size="20"> <font color="#FF0000">请填写正确的信箱</font></td>
    </tr>
    <tr> 
      <td width="100" height="25">显示方式：</td>
      <td colspan="2">　 
        <input type="radio" value="0" <%if rs("img")=0 then Response.Write "checked"%>  name="img">文字 
        <input type="radio" value="1" <%if rs("img")=1 then Response.Write "checked"%> name="img">
        图标 　<font color="#FF0000">选择图标时请注意填写图标地址<%If linklogo=1 then%>或上传图标 <%End if%>！</font></td>
    </tr>

    <tr> 
      <td width="100" height="25">网站LOGO</td>
      <td colspan="2"><input name="logo" type="text" class="input" id="logo" size="40" maxlength="255" value="<%=trim(rs("logo"))%>">没有请留空 <%If rs("logo")<>"0" then%><img src="../<%=rs("logo")%>" border=1 width="88"  height="31"><%End if%><br><%If linklogo=1 then%><iframe name="ad" frameborder=0 width="450" height="35" scrolling=no src="../DNJY_upload.asp?ttup=link"></iframe><br>&nbsp;&nbsp;<font color="#ff0000">允许上传的文件类型：gif|jpgs|jpg|bmp&nbsp;&nbsp;&nbsp;仅jpg格式图片可幻灯显示!&nbsp;&nbsp;限：88*31像素，200k以下</font><%End if%> </td>
    </tr>
   <tr> 
      <td width="100" height="25">网站说明</td>
      <td colspan="2"><textarea name="info" cols="50" rows="3" id="info" onkeyUp="textLimitCheck(this, 50);"><%=trim(rs("info"))%></textarea>&nbsp;&nbsp; <br>限 50 个字符 已输入 <font color="#CC0000"><span id="messageCount">0</span></font> 个字></td>
    </tr>

  <tr bgcolor="#eeeeee">
    <td height="35" colspan="10"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><div align="center">
		
            <input type="submit" value="确认" name="B1" onClick="if(wzname.value=='' || url.value==''){alert('名称和电话都不能为空!');return false;}">
        </div></td>
        <td><div align="center">
            <input type="reset" value="取消" name="B1">
        </div></td>
      </tr>          
        </FORM>
        <%Call CloseRs(rs) %>
      </TABLE>
	<BR></TD>
  </TR>
</TABLE>
<%End if%>
<%If adminlink="2" Then
 If addlink=1 Then
 response.write "本站不开放友情链接申请,请联系管理员。信箱："&mymail&" QQ："&qq&""
 response.End
 End if%>
<FORM name="form" method="post" action="?action=add" language="javascript" onsubmit="return CheckForm()">
<TABLE width="600" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
  <TR> 
    <TD height="20" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">申 请 友 情 链 接<br>（友情链接要经本站审核）</FONT></TD>
  </TR>
  <TR> 
    <TD bgcolor="#FFFFFF">
	<BR>
	<TABLE width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
   <TR> 
    <TD bgcolor="#FFFFFF"> <BR>
	<TABLE width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
            <tr> 
      <td width="100" height="25">所属地区：</td>
      <td colspan="2">
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
		//dim count4:
		count4 = 0
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
    document.form.city_two.length = 0; 
	document.form.city_two.options[0] = new Option('选择城市','');
	document.form.city_three.length = 0; 
	document.form.city_three.options[0] = new Option('选择城市','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.form.city_two.options[document.form.city_two.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
	
	function changelocation4(locationid,locationid1)
    {
    document.form.city_three.length = 0; 
    document.form.city_three.options[0] = new Option('选择城市','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount4; i++)
        {
            if (subcat4[i][2] == locationid)
            { 
			if (subcat4[i][1] == locationid1)
			{
                document.form.city_three.options[document.form.city_three.length] = new Option(subcat4[i][0], subcat4[i][3]);
            }        
        }
       } 
    }            
              </SCRIPT>
          <SELECT name="city_one" size="1" id="select" onChange="changelocation(document.form.city_one.options[document.form.city_one.selectedIndex].value)">
            <OPTION value="" selected>选择地区</OPTION>
            <%set rs=conn.execute("select * from DNJY_city where id>0 and twoid=0 order by indexid asc")
if rs.eof or rs.bof then
response.write "<option value=''>没有分类</option>"
else
do until rs.eof
response.write "<option value='"&rs("id")&"'>"&rs("city")&"</option>"
    rs.movenext
    loop
end if
Call CloseRs(rs)
%>
          </SELECT>
          <SELECT name="city_two" id="select6" onChange="changelocation4(document.form.city_two.options[document.form.city_two.selectedIndex].value,document.form.city_one.options[document.form.city_one.selectedIndex].value)">
            <OPTION value="" selected>选择城市</OPTION>
          </SELECT>
          <SELECT name="city_three" id="select7">
            <OPTION value="" selected>选择城市</OPTION>
          </SELECT></td>
    </tr>  
    <tr> 
      <td width="100" height="25"><font color="#FF0000">*</font>网站名称</td>
      <td colspan="2">
<INPUT name="wzname" type="text" id="wzname" value="" size="20"  maxlength="16" onpropertychange="kn()"> <span id=num>你还可以输入16个字符</span></td>
    </tr>
    <tr> 
      <td width="100" height="25"><font color="#FF0000">*</font>网址</td>
      <td colspan="2">
    <INPUT name="url" type="text" id="url" value="" size="40"></td>
    </tr>
    <tr> 
      <td width="100" height="25"><font color="#FF0000">*</font>管理用户名</td>
      <td colspan="2">
<INPUT name="name" type="text" id="name" value="" size="20"  maxlength="8" onpropertychange="kn2()"> <span id=num2>你还可以输入8个字符</span></td>
    </tr>
    <tr> 
      <td width="100" height="25"><font color="#FF0000">*</font>管理密码</td>
      <td colspan="2">
<INPUT name="pass" type="text" id="pass" value="" size="20"  maxlength="8" onpropertychange="kn3()"> <span id=num3>你还可以输入8个字符</span></td>
    </tr>
    <tr> 
      <td width="100" height="25"><font color="#FF0000">*</font>密码确认</td>
      <td colspan="2">
<INPUT name="passs" type="text" id="passs" value="" size="20"></td>
    </tr>
    <tr> 
      <td width="100" height="25"><font color="#FF0000">*</font>电子信箱</td>
      <td colspan="2">
<INPUT name="mail" type="text" id="mail" value="" size="20"> <font color="#FF0000">请填写正确的信箱</font></td>
    </tr>
    <tr> 
      <td height="25">显示方式：</td>
      <td colspan="2"> 
        <input type="radio" name="img" value="0" checked>
        文字 
        <input name="img" type="radio" value="1" >
        图标 　<font color="#FF0000">选择图标时请注意填写图标地址<%If linklogo=1 then%>或上传图标<%end if%>！</font></td>
    </tr> 

    <tr> 
      <td width="100" height="25">网站LOGO</td>
      <td colspan="2"><input name="logo" type="text" class="input" id="logo" size="40" maxlength="255" value="">填写图片链接地址，没有请留空<%If linklogo=1 then%>或上传LOGO<br><iframe name="logo" frameborder=0 width="450" height="35" scrolling=no src="../DNJY_upload.asp?ttup=link"></iframe><br>&nbsp;&nbsp;<font color="#ff0000">允许上传的文件类型：gif|jpgs|jpg|bmp&nbsp;&nbsp;&nbsp;仅jpg格式图片可幻灯显示!&nbsp;&nbsp;限：88*31像素，200k以下</font> <%end if%></td>
    </tr>
    <tr> 
      <td width="100" height="25">网站说明</td>
      <td colspan="2"><textarea name="info" cols="50" rows="3" id="info" onkeyUp="textLimitCheck(this, 50);"></textarea>&nbsp;&nbsp; <br>限 50 个字符 已输入 <font color="#CC0000"><span id="messageCount">0</span></font> 个字></td>
    </tr>
  <tr bgcolor="#eeeeee">
    <td height="35" colspan="10"><table width="70%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><div align="center">
		
            <input type="submit" value="确认" name="B1" >
        </div></td>
        <td><div align="center">
            <input type="reset" value="取消" name="B1">
        </div></td>
      </tr>          
        </FORM>
       
     
      </TABLE>
	<BR></TD>
  </TR>
</TABLE>
<%End if%>
<TABLE width="600" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
  <TR> 
    <TD  align="center" bgcolor="#ff9AE1"><FONT color="#FFFFFF" style="font-size:14px">&nbsp;
    </FONT><FONT color="#FFFFFF" style="font-size:14px">链接本站代码<p>申请链接后,请及时在贵站加上链接,以免影响审核</FONT></TD>
	<TR> 
    <TD bgcolor="#ffffff"><b>网站名称:</b> <%=webname%><p>
	<b>网站地址:</b> http://<%=web%><p>
	<b>LOGO地址:</b> http://<%=web%>/images/logo.gif &nbsp;&nbsp;&nbsp;<img src="http://<%=web%>/images/logo.gif" border=1 width="88"  height="31"><p>
	<b>网站说明:</b> <%=webname%>
	<p></TD>
  </TR>
</TABLE>
</BODY> 
</HTML> 
<%Call CloseDB(conn)%>