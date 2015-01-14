<!--#include file="conn.asp"-->
<!--#include file=top.asp-->
<!--#include file="Include/err.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
if not ChkPost() then 
response.write "禁止站外提交或访问"
response.end 
end if
if lmkg2="1" then
call errdick(50)
response.end
end If

if addip<>"0" then
if ip<>"" then
'ip=checktext(ip)
addip=split(cstr(ip),"@")
for N=0 to UBound(addip)
if instr(Request.serverVariables("REMOTE_ADDR"),addip(n))>0 then
errdick(43)
response.end
end if
next
end if
end If

if addxinxi=0 then
response.redirect "login.asp"
end if
%>
<%'禁止网页缓存
'Response.Buffer = True 
'Response.ExpiresAbsolute = Now() - 1 
'Response.Expires = 0 
'Response.CacheControl = "no-cache"
%>
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">

<!--文字验证码调用开始-->
<SCRIPT LANGUAGE=javascript>
/*显示认证码 o start1*/
function get_Code() {
        var Dv_CodeFile = "Dv_GetCode.asp";
        if(document.getElementById("imgid"))
                document.getElementById("imgid").innerHTML = '<img src="'+Dv_CodeFile+'?t='+Math.random()+'" alt="点击刷新验证码" style="cursor:pointer;border:0;vertical-align:middle;height:30px;" onclick="this.src=\''+Dv_CodeFile+'?t=\'+Math.random()" />'
}
/*o end*/
</script>
<script language="JavaScript" type="text/javascript">
var dvajax_request_type = "GET";
</script>
<script language="JavaScript" src="dv_ajax.js" type="text/javascript"></script>
<script language=javascript src=Include/mouse_on_title.js></script>
<!---文字验证码调用结束-->
<script language="javascript">
<!--

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

//标题字节数限制，与下面配合=========
String.prototype.Length=function(){
temp=this.replace(/([^\x00-\xff])/g,"$1$1");
return temp.length
}
//显示输入字符数

function kn()
  {
   var ln=myform.biaoti.value.Length();
     num.innerText='您还可以输入:'+(<%=biaotinum%>-ln)+'个字符'+(<%=biaotinum/2%>-ln/2)+'汉字';
      //if (ln>=8) form.biaoti.readOnly=true;  //这行是如果满足条件表单无法输入和修改
}

//说明字节数限制，与下面配合=========
String.prototype.Length=function(){
temp=this.replace(/([^\x00-\xff])/g,"$1$1");
return temp.length
}
//显示输入字符数

function kn2()
  {
   var ln=myform.memo.value.Length();
     num2.innerText='您还可以输入:'+(<%=memonum%>-ln)+'个字符'+(<%=memonum/2%>-ln/2)+'汉字';
      //if (ln>=8) form.memo.readOnly=true;  //这行是如果满足条件表单无法输入和修改
}

function myform_onsubmit() {
if (document.myform.city_one.value==0)
	{
	  alert("请选择地区！")
	  document.myform.city_one.focus()
	  return false
	 }


if (document.myform.type_one.value==0)
	{
	  alert("请选择信息一级分类！")
	  document.myform.type_one.focus()
	  return false
	 }

if ((document.myform.biaoti.value.Length()><%=biaotinum%>) || (document.myform.biaoti.value.Length()<4)) //字节数限制，与上面配合
     {
	 alert("信息标题长度要求4－<%=biaotinum%>字节(<%=biaotinum/2%>汉字)，最大不能超100字节,请重新输入！");
	  document.myform.biaoti.focus()
	  return false
  }
if (document.myform.keywords.value=="" )
	{
	  alert("请填信息关键词！")
	  document.myform.keywords.focus()
	  return false
	 }
if (document.myform.keywords.value.length>100)
	{
	  alert("关键词不能超过100个字符")
	  document.myform.keywords.focus()
	  return false
	 }	 
if (document.myform.leixing.value=="" )
	{
	  alert("请选择交易类型！")
	  document.myform.leixing.focus()
	  return false
	 }	 
//if (document.myform.scjiage.value=="" )
//	{
//	  alert("请填市场价格！")
//	  document.myform.scjiage.focus()
//	  return false
//	 }
//if (isNaN(document.myform.scjiage.value))
//	{
//	  alert("市场价格应为数字！")
//	  document.myform.scjiage.focus()
//	  return false
//	 }	
if (document.myform.jiage.value=="" )
	{
	  alert("请填交易价格！")
	  document.myform.jiage.focus()
	  return false
	 }		 
if (isNaN(document.myform.jiage.value))
	{
	  alert("交易价格应为数字！")
	  document.myform.jiage.focus()
	  return false
	 }		 		

if (document.myform.memo.value=="" )
	{
	  alert("请填信息说明！")
	  document.myform.memo.focus()
	  return false
	 }
if (document.myform.memo.value.Length()><%=memonum%>)  //字节数限制，与上面配合
     {
	 alert("信息说明长度要求<%=memonum%>字节(<%=memonum/2%>汉字)，请重新输入！");
	  document.myform.memo.focus()
	  return false
  }
if (document.myform.name.value=="" )
	{
	  alert("请填联系人！")
	  document.myform.name.focus()
	  return false
	 }

if(document.myform.dianhua.value=="" && document.myform.CompPhone.value=="") 
	{
	    alert("固定电话和移动电话不能同时为空!");
		document.myform.dianhua.focus()
	    return false;
	}
//var sMobile = document.myform.dianhua.value
//if((document.myform.dianhua.value!="") && (!(/^([0\+]\d{2,3}-)?(0[0-9]{2,3}\-)?([2-9][0-9]{6,7})+(\-[0-9]{1,4})?$/.test(sMobile)))) //(国家代号-区号-电话-分机格式(可选)：+086-010-85858585-2538)
//	{
//		alert("不规范的电话号码！");
//		document.myform.dianhua.focus();
//		return false;
//	}	
//var sMobile = document.myform.CompPhone.value
//if((document.myform.CompPhone.value!="") && (!(/^((\(\d{3}\))|(\d{3}\-))?13[0-9]\d{8}|15[0-9]\d{8}|18[0-9]\d{8}/.test(sMobile)))) //(检验13，15，18号段)
//	{
//		alert("不是完整的11位手机号或者正确的手机号前七位！");
//		document.myform.CompPhone.focus();
//		return false;
//	}
    if((!checkemail(myform.email.value))&&(document.myform.email.value!=""))
	{
    alert("您输入Email地址不正确，请重新输入!");
    document.myform.email.focus();
    return false;
    }
   if(document.myform.email.value.length>30 )
	{
	    alert("信箱不能超过30个字符!");
	    document.myform.email.focus();
	    return false;
	}		
<%if codekg1=1 then%>
    if (document.myform.wenti.value=="" )
	{	  
      alert("验证答案不能为空，请重新输入！");
	  document.myform.wenti.focus()
	  return false
	 }
	
    if(document.myform.wenti.value != document.myform.daan.value) 
	{	  
      alert("验证答案不对，请重新输入！");
	  document.myform.wenti.focus()
	  return false
	 }
 <%end if%>
 <%if codekg2=1 then%>
    if (document.myform.yzm.value=="" )
	{	  
      alert("数字验证码不能为空，请重新输入！");
	  document.myform.yzm.focus()
	  return false
	 }
<%end if%>
 <%if codekg3=1 then%>
    if (document.myform.code.value=="" )
	{	  
      alert("文字验证码不能为空，请重新输入！");
	  document.myform.code.focus()
	  return false
	 }
<%end if%>
<%if codekg4=1 then%>
    if (document.myform.ttdv.value=="" )
	{	  
      alert("验证星期不能为空，请重新输入！");
	  document.myform.ttdv.focus()
	  return false
	 }
<%end if%>
<%if codekg5=1 then%>
    if (document.myform.validate_codee.value=="" )
	{	  
      alert("算式验证码不能为空，请重新输入！");
	  document.myform.validate_codee.focus()
	  return false
	 }
<%end if%>
	  
}
// --></script>
</head>

<body topmargin="0" leftmargin="0">

<div align="center">

  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="980" height="428">
    <tr>
      <td width="215" height="356" valign="top"><div align="center">
      <!--#include file=left.asp--></div></td>
	  <td width="5">&nbsp;</td>
      <td width="279" height="356" valign="top" bgcolor="#FFFFFF"><table width="760" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="dq1">

        <tr>
          <td width="760" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" align="right" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse">
          <form method="POST" name="myform" id="myform" language="javascript" onSubmit="return myform_onsubmit()" action="addtourists_save.asp?<%=C_ID%>">
             
                <tr>
                  <td colspan="3"><div align="center">
                      <table width="100%" height="100%" border="0" align="center" cellpadding="6" cellspacing="0">
                       
                        <tr bgcolor="#FACA38">
                          <td height="1" colspan="2" ><font color="#FF0000">&nbsp;&nbsp;&nbsp;&nbsp;声明：在<u><%=title%></u>中发布信息必须对您的言行负责，遵守国家有关法律、法规，遵守道德和诚信原则；禁止利用本站进行危害国家和欺诈他人的行为，禁止发布违禁物品、伪劣商品及虚假交易信息；本站保留删除违规信息和向有关部门举报的权利；信息言论若违反规定而受到法律追究者，责任由发布者自负。对群发垃圾信息坚决删除！</font></td>
                        </tr>
						<TABLE>
                        <TR>
							<TD><TABLE> 
							<tr>
                          <td height="25" width="100%" colspan="2" align="center"><b>注册会员可享受置顶、上传图片、标题变色、验证道具、详细说明所见所得编辑器和<br>发布竞价信息等权限，联络信息自动填入，以后发布更便捷，建议<a href="<%=reg%>"><font color="#0000FF">注册</font></a>后发布信息！</b><br>带<font color="#ff0000">*</font>为必填&nbsp; </td>     
                        </tr>
                        <tr>
                          <td height="25" width="60"> 交易地区：</td>
                          <td height="25" width="700">
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
    document.myform.city_two.length = 0; 
	document.myform.city_two.options[0] = new Option('选择城市','');
	document.myform.city_three.length = 0; 
	document.myform.city_three.options[0] = new Option('选择城市','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.myform.city_two.options[document.myform.city_two.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
	
	function changelocation4(locationid,locationid1)
    {
    document.myform.city_three.length = 0; 
    document.myform.city_three.options[0] = new Option('选择城市','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount4; i++)
        {
            if (subcat4[i][2] == locationid)
            { 
			if (subcat4[i][1] == locationid1)
			{
                document.myform.city_three.options[document.myform.city_three.length] = new Option(subcat4[i][0], subcat4[i][3]);
            }        
        }
       } 
    }            
              </SCRIPT>
          <SELECT name="city_one" size="1" id="select" onChange="changelocation(document.myform.city_one.options[document.myform.city_one.selectedIndex].value)">
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
          <SELECT name="city_two" id="select6" onChange="changelocation4(document.myform.city_two.options[document.myform.city_two.selectedIndex].value,document.myform.city_one.options[document.myform.city_one.selectedIndex].value)">
            <OPTION value="" selected>选择城市</OPTION>
          </SELECT>
          <SELECT name="city_three" id="select7">
            <OPTION value="" selected>选择城市</OPTION>
          </SELECT>

<font color="#ff0000"> *</font></td>
                        </tr>
						<tr>
                          <td height="25" width="60"> 信息类别：</td>
                          <td height="25" width="450"><%
set rs=conn.execute("select * from DNJY_type where id>0 and twoid>0 and threeid=0 order by indexid")
%>
          <SCRIPT language = "JavaScript">
var onecount2;
onecount2=0;
subcat2 = new Array();
        <%
		dim count2:count2 = 0
        do while not rs.eof 
        %>
subcat2[<%=count2%>] = new Array("<%=rs("name")%>","<%=rs("id")%>","<%=rs("twoid")%>");
        <%
        count2 = count2 + 1
        rs.movenext
        loop
Call CloseRs(rs)
        %>
onecount2=<%=count2%>;
          </SCRIPT>
          <%
set rs=conn.execute("select * from DNJY_type where id>0 and twoid>0 and threeid>0 order by indexid")
%>
          <SCRIPT language = "JavaScript">
var onecount3;
onecount3=0;
subcat3 = new Array();
        <%
		dim count3:count3 = 0
        do while not rs.eof 
        %>
subcat3[<%=count3%>] = new Array("<%=rs("name")%>","<%=rs("id")%>","<%=rs("twoid")%>","<%=rs("threeid")%>");
        <%
        count3 = count3 + 1
        rs.movenext
        loop
Call CloseRs(rs)
        %>
onecount3=<%=count3%>;



function changelocation2(locationid)
    {
    document.myform.type_two.length = 0; 
    document.myform.type_two.options[0] = new Option('二级分类','');
	document.myform.type_three.length = 0; 
    document.myform.type_three.options[0] = new Option('三级分类','');
    var locationid=locationid;
	var i;
    for (i=0;i < onecount2; i++)
        {
            if (subcat2[i][1] == locationid)
            { 
                document.myform.type_two.options[document.myform.type_two.length] = new Option(subcat2[i][0], subcat2[i][2]);
            }        
        }
       
    }    
     function changelocation3(locationid,locationid1)
    {
    document.myform.type_three.length = 0; 
    document.myform.type_three.options[0] = new Option('三级分类','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount3; i++)
        {
            if (subcat3[i][2] == locationid)
            { 
			if (subcat3[i][1] == locationid1)
			{
                document.myform.type_three.options[document.myform.type_three.length] = new Option(subcat3[i][0], subcat3[i][3]);
            }        
        }
       } 
    }            
	      </SCRIPT>
          <SELECT name="type_one" size="1" id="select8" onChange="changelocation2(document.myform.type_one.options[document.myform.type_one.selectedIndex].value)">
            <OPTION value="" selected>一级分类</OPTION>
            <%set rs=conn.execute("select * from DNJY_type where id>0 and twoid=0 order by indexid asc")
if rs.eof or rs.bof then
response.write "<option value=''>没有分类</option>"
else
do until rs.eof
response.write "<option value='"&rs("id")&"'>"&rs("name")&"</option>"
    rs.movenext
    loop
end if
Call CloseRs(rs)
%>
          </SELECT>
          <SELECT name="type_two" id="select9" onChange="changelocation3(document.myform.type_two.options[document.myform.type_two.selectedIndex].value,document.myform.type_one.options[document.myform.type_one.selectedIndex].value)">
            <OPTION value="" selected>二级分类</OPTION>
          </SELECT>
          <SELECT name="type_three" id="select10">
            <OPTION value="" selected>三级分类</OPTION>
          </SELECT> <font color="#ff0000"> *</font></td>
                        </tr>
                        <tr>
                          <td height="25" width="60"> 信息标题：<!---本站能识别群发垃圾信息，对群发垃圾信息坚决删除！---></td>
                          <td height="25" width="450"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" maxlength="100" name="biaoti" size="60" onpropertychange="kn()" value="">
                            <font color="#ff0000"> *</font><br>（<font color="#CC5200"><%=biaotinum%>字节，<%=biaotinum/2%>汉字以内</font>）<span id=num>你还可以输入<%=biaotinum%>个字节，<%=biaotinum/2%>汉字</span></td>
                        </tr>
                        <tr>
                          <td height="25" width="60"> 关 键 词：</td>
                          <td height="25" width="450"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" maxlength="100" name="keywords" size="60" value="">
                            <font color="#ff0000"> *</font> <input type="button" name="btn" value="从标题复制" title="复制标题到关键词" onClick="CopyWebbiaoti(document.myform.biaoti.value);"><br>（<font color="#CC5200">信息关键词，用“,”号分隔，限100字节内。</font>）</td>
                        </tr>
                        <tr>
                          <td height="25" width="71"> 交易类型：</td>
                          <td height="25" width="408"><%call leixs()%>  <font color="#ff0000"> *</font></td>
                        </tr>
                        <!--<tr>
                          <td height="25" width="60"> 市场价格：</td>
                          <td height="25" width="450">
                              <input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="scjiage" size="10" maxlength="30" value="" onKeyUp="if(isNaN(value)){alert('价格只允许输入数字');value='';}"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))">
                            元 <font color="#ff0000"> *</font>（0<font color="#CC5200">元表示面议</font>）</td>
                        </tr>-->
                        <tr>
                          <td height="25" width="60"> 交易价格：</td>
                          <td height="25" width="450">
                              <input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="jiage" size="10" maxlength="30" value="" onKeyUp="if(isNaN(value)){alert('价格只允许输入数字');value='';}"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))">
                            元 <font color="#ff0000"> *</font>（0<font color="#CC5200">元表示面议</font>）</td>
                        </tr>
                        <tr>
                          <td height="25">发布天数：</td>
                          <td height="25"><select name="days" style="BACKGROUND-COLOR: #FFCCFF; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;">
						  <option value="1">一天</option>
						  <option value="2">二天</option>
						  <option value="3">三天</option>
						  <option value="4">四天</option>
						  <option value="5">五天</option>
						  <option value="6">六天</option>
                          <option value="7">一星期</option>
						  <option value="14">二星期</option>
						  <option value="21">三星期</option>
                          <option value="30">一个月</option>
						  <option value="60">二个月</option>
                          <option value="90">三个月</option>
					      <option value="180">半年</option>
						  <option value="365">一年</option>
                          <option value="1100">三年</option>
                            </select>
                              </td>
                        </tr>
					    <tr>
                        <td height="25">阅读权限：</td>
                        <td width="400" height="25"><input type="radio" name="Readinfo" value="0" checked>游客	<input type="radio" name="Readinfo" value="1">普通会员	<input type="radio" name="Readinfo" value="2">VIP会员	
                             <br>（限制信息详细说明及联络方式，会员级别高向低兼容）</td>
                      </tr>						
					    <!--<tr>
                        <td height="25">阅读权限：</td>
                        <td width="400" height="25"><input type="radio" name="Readinfo" value="1" checked>普通会员	<input type="radio" name="Readinfo" value="2">VIP会员	
                             <br>（限制信息详细说明及联络方式，会员级别高向低兼容）</td>
                      </tr>
					    <tr>
                        <td height="25">外观颜色：</td>
                        <td height="25" width="700"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="wcolor" size="23" maxlength="50" value=""> <font color="#ff0000"> *</font> &lt;=【<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="wcolor.value='黑色'" 
      color=#993300>黑色</FONT></FONT>】【<FONT 
      color=blue><FONT style="CURSOR: hand" onclick='wcolor.value="红色"' 
      color=#993300>红色</FONT></FONT>】【<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="wcolor.value='白色'" 
      color=#993300>白色</FONT></FONT>】【<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="wcolor.value='银色'" 
      color=#993300>银色</FONT></FONT>】【<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="wcolor.value='棕色'" 
      color=#993300>棕色</FONT></FONT>】</td>
                      </tr>
					    <tr>
                        <td height="25">内饰颜色：</td>
                        <td height="25" width="700"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="ncolor" size="23" maxlength="50" value=""> <font color="#ff0000"> *</font> &lt;=【<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="ncolor.value='黑色'" 
      color=#993300>黑色</FONT></FONT>】【<FONT 
      color=blue><FONT style="CURSOR: hand" onclick='ncolor.value="灰色"' 
      color=#993300>灰色</FONT></FONT>】【<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="ncolor.value='棕色'" 
      color=#993300>棕色</FONT></FONT>】【<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="ncolor.value='米色'" 
      color=#993300>米色</FONT></FONT>】【<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="ncolor.value='紫色'" 
      color=#993300>紫色</FONT></FONT>】</td>
                      </tr>--->
</TABLE></TD>
							
                        </TR>
                        </TABLE>

<TABLE>
<TR>
	<TD><TABLE> 
	<tr>
                          <td height="25" width="60"><p>详细说明：<br><%If Filterhtm>0 Then%>(游客发布信息屏蔽Html代码，为确保内容不乱，可将从网页复制的内容粘贴在记事本，再复制过来)<%End if%></p></td>
                          <td height="25" width="700">
						  <textarea rows="18" name="memo"  cols="90" onpropertychange="kn2()"><%=request("memo")%></textarea><font color="#ff0000"> *</font><br>（<font color="#CC5200"><%=memonum%>字节，<%=memonum/2%>汉字以内</font>）<span id=num2>您还可以输入<%=memonum%>个字节，<%=memonum/2%>汉字</span><br>（注册为会员可享受HTML编辑器，所见所得，显示更美观）</td>
                        </tr>
                        <tr>
                          <td height="25" width="60"> 真实姓名：</td>
                          <td height="25" width="700"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="name" size="23" maxlength="15" value="">
                            &nbsp; <font color="#ff0000"> *</font> </td>
                        </tr>
                         <tr>
                          <td height="25">移动电话：</td>
                          <td width="408" height="25"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="CompPhone" size="23" maxlength="20" value="" onKeyUp="if(isNaN(value)){alert('移动电话只允许输入数字');value='';}" onafterpaste="this.value=this.value.replace(/\D/g,'')">  &nbsp; <font color="#0000ff"> *</font>（<font color="#CC5200">固定电话和移动电话必填其一</font>）</td>
                        </tr>

                        <tr>
                          <td height="25">固定电话：</td>
                          <td width="600" height="25"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="dianhua" size="23" maxlength="20" value="">
                            &nbsp; <font color="#0000ff"> *</font>（<font color="#CC5200">固定电话和移动电话必填其一</font>） <br>(国家代号-区号-电话-分机格式(可只填区号+电话或电话)：+086-010-85858585-2538)</td>
                        </tr>
                        <tr>
                          <td height="25">传真号码：</td>
                          <td width="600" height="25"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="fax" size="23" maxlength="20" value="">
                            &nbsp; (国家代号-区号-电话-分机格式(可只填区号+电话或电话)：+086-010-85858585-2538)</td>
                        </tr>
                        <tr>
                          <td height="25" width="71">电子信箱：</td>                          
                          <td width="520" height="-5"><input name="email" type="text" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" value="" size="23" maxlength="40" >
                            &nbsp; （<font color="#FF0000">重要：</font><font color="#CC5200">该邮件地址将接收你的交易信息，请正确填写！</font>）</td>
                        </tr>
                        <tr>
                          <td height="25">QQ 号 码：</td>
                          <td width="408" height="-1"><input name="qq" type="text" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;"  size="23" maxlength="20" onKeyUp="if(isNaN(value)){alert('QQ号码只允许输入数字');value='';}" onafterpaste="this.value=this.value.replace(/\D/g,'')">
                            &nbsp;</td>
                        </tr>
                        <tr>
                          <td height="25"> 微信号码：</td>
                          <td height="25" width="408"><input name="weixin" type="text" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" size="23" maxlength="50" value="">
                            &nbsp; </td><!--123:-->
                        </tr>						
                        <tr>
                          <td height="25"> 邮政编码：</td>
                          <td height="25" width="408"><input name="youbian" type="text" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" size="23" maxlength="30" value="" onKeyUp="if(isNaN(value)){alert('邮政编码只允许输入数字');value='';}" onafterpaste="this.value=this.value.replace(/\D/g,'')">
                            &nbsp; </td>
                        </tr>
                        <tr>
                          <td height="25"> 公司名称：</td>
                          <td height="25" width="408"><input name="mpname" type="text" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" size="50" maxlength="50" value="">
                            &nbsp; </td><!--123:-->
                        </tr>						
                        <tr>
                          <td height="25"> 详细地址：</td>
                          <td height="25" width="408"><input name="dizhi" type="text" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" size="50" maxlength="100" value="">
                            &nbsp; </td><!--123:-->
                        </tr>
                        <tr>
                          <td height="25"> 删除密码：</td>
                          <td height="25" width="408"><input name="delpas" type="text" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" size="23" maxlength="30" value="">
                            &nbsp;（游客删除本条信息专用，请牢记）</td>
                        </tr>
<%if codekg1=1 Or codekg2=1 Or codekg3=1 Or codekg4=1 Or codekg5=1 then%>
                       <tr>
<td height="25" colspan="2">=======================<b>为防止群发垃圾信息，请耐心填写下列验证码</b>=======================</td>
                       </tr>
 <%End if%>                 
<%if codekg1=1 then
	  '问答式验证
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
	        <tr >
        <td >问答验证：</td>
        <td>问题：<%=rscheck("wenti")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;答案：<input type="text" name="wenti" size="12">  <font color="#ff0000"> *</font></td>
      </tr><input type="hidden" name="daan" value="<%=rscheck("daan")%>">
	<%rscheck.close
	set rscheck=Nothing	
	End If
if codekg2=1  then%>
 <tr >
        <td >数字验证：</td>
        <td><input type="text" name="yzm" size="4" /><img src="tt_getcode.asp" alt="验证码,看不清楚?请点击刷新验证码" height="18" style="cursor : pointer;" onClick="this.src='tt_getcode.asp?t='+(new Date().getTime());" />&nbsp;&nbsp; <font color="#ff0000"> *</font> （验证码,看不清楚?请点击刷新验证码）</td>
      </tr>
	  <%End if
	if codekg3=1 then%>
      <tr >
        <td >文字验证：</td>
        <td><input type="text" name="code" id="code" size="4" maxlength="4" tabindex="6" onFocus="get_Code();this.onfocus=null;" onKeyUp="dv_ajaxcheck('checke_dvcode','code');"  autocomplete="off"/> <font color="#ff0000"> *</font>
    <span id="imgid" style="color:red">点击获取验证码</span><span id="isok_code"></span></td>
      </tr>
<%End if%>
<%if codekg4=1 then%>

<tr>                                                
  <td height="25"> 验证星期：</td>
      <td height="25" width="408">今天是 <select name="ttdv">
<option value="" selected="selected">请选择</option>
<option value="1">星期一</option>
<option value="2">星期二</option>
<option value="3">星期三</option>
<option value="4">星期四</option>
<option value="5">星期五</option>
<option value="6">星期六</option>
<option value="0">星期日</option>
</select>
       <font color="#ff0000"> *</font></td>
  </tr>  
<%End if%>
<%if codekg5=1 then%>
<tr>                                                
  <td height="25"> 算式验证：</td>
      <td height="25" width="408"><img src="tt_getcodee.asp" alt="验证码,看不清楚?请点击刷新验证码" height="10" style="cursor : pointer;" onClick="this.src='tt_getcodee.asp?t='+(new Date().getTime());" /><input name="validate_codee" type="text" tabindex="3" value="" size="4" maxlength="4" onKeyUp="value=value.replace(/[^\d|^\.]/g,'')"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))"/> &nbsp;&nbsp;<font color="#ff0000"> *</font> （验证码,看不清楚?请点击刷新验证码）
       </td>
    </tr>
<%End if%>
  <tr bgcolor="#DBDBDB">
                          <td height="1" colspan="2"></td>
                        </tr>
                        <tr>
                          <td height="30" colspan="2"><div align="center" style="color: #FF0000">
                              <div align="left" style="color: #FF0000"><% if xinxish=0 or onoff=0 then %><span style="font-weight: bold">声明：</span>游客快速发布需要等管理员审核才能发表，请见谅。<%Else%><span style="font-weight: bold">声明：</span>当前系统设置快速发布无需等管理员验证便可直接发布。</div></div><%end if%></td>
                        </tr>
                        <tr bgcolor="#D9D9D9">
                          <td height="1" colspan="2"></td>
                        </tr>
                        <tr>
                          <td height="10" colspan="2"></td>
                        </tr>
                        <tr>
                          <td height="25" colspan="2" align="center"><table width="65%"  border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td><div align="center">
                                    <input class="inputb" name="I2" type="submit" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #a2a2a2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #a2a2a2 1px solid; BORDER-TOP: #ffffff 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;"  value="提交发布" border="0" /></font>&nbsp;&nbsp;（<font color="#ff0000">请检查您所填内容是否准确</font>） 
                                </div></td>
                                <td><div align="center">
                                    <input class="inputb" name="I2" type="reset" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #a2a2a2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #a2a2a2 1px solid; BORDER-TOP: #ffffff 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" onClick="javascript:return CheckForm();" value="取消发布" border="0" />
                                </div></td>
                              </tr>
                          </table></td> </form>
</TABLE></TD>
</TR>
</TABLE>

                        </tr>
                      </table>
                  </div></td>
                </tr>               
             
          </table></td>
        </tr>
      </table>
        <div align="center">
        <center>
        </center>
        </div>        </td>
      <td width="4" valign="top" bgcolor="#E6E6E6"></td>
    </tr>
    <tr>
      <td height="26" colspan="3"><!--#include file=footer.asp--></td>
    </tr>
  </table>
  </center>
</div>

</body>
</html>
