<!--#include file="conn.asp"-->
<!--#include file=top.asp-->
<!--#include file="Include/err.asp"-->
<!--#include file=usercookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Dim Amount
username=request.cookies("dick")("username")
Call OpenConn
set rs=server.createobject("adodb.recordset")
sql = "select name,vip,Amount,email,dianhua,a,b,c,d from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,1
vip=rs("vip")
Amount=rs("Amount")
if rs.eof then
errdick(9)
response.end
end if
%>

<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
<script language=javascript src=Include/mouse_on_title.js></script>
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

//说明字节数限制，与下面配合=========
String.prototype.Length=function(){
temp=this.replace(/([^\x00-\xff])/g,"$1$1");
return temp.length
}

function formcheck(){
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


if (document.myform.biaoti.value=="" )
	{
	  alert("请填信息标题！")
	  document.myform.biaoti.focus()
	  return false
	 }	 
if (document.myform.biaoti.value.length>100)
	{
	  alert("标题不能超过100个字符！")
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
	  alert("关键词不能超过100个字节！")
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
var editor=KindEditor.create('#editor');
if (editor.isEmpty())
	{
	  alert("信息说明不能为空！")	  
	  return false
	 }
if (document.myform.memo.value.Length()><%=memonum%>)  //字节数限制，与上面配合
     {
	 alert("信息说明长度要求<%=memonum%>字节(<%=memonum/2%>汉字)，请重新输入！");
	  return false
  }	 
if (document.myform.name.value=="" )
	{
	  alert("请填真实姓名！")
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
//	var sMobile = document.myform.CompPhone.value
//	if((document.myform.CompPhone.value!="") && (!(/^((\(\d{3}\))|(\d{3}\-))?13[0-9]\d{8}|15[0-9]\d{8}|18[0-9]\d{8}/.test(sMobile)))) //(检验13，15，18号段)
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
    if (document.myform.hfje.value=="" )
	{	  
      alert("竞价金额不能为空，请重新输入或保留为0！");
	  document.myform.hfje.focus()
	  return false
	 }	
	  
}
// --></script>
<!--kindeditor编辑器-->
	<link rel="stylesheet" href="../KindEditor/themes/default/default.css" />
	<link rel="stylesheet" href="../KindEditor/plugins/code/prettify.css" />
	<script charset="utf-8" src="../KindEditor/kindeditor.js"></script>
	<script charset="utf-8" src="../KindEditor/lang/zh_CN.js"></script>
	<script charset="utf-8" src="../KindEditor/plugins/code/prettify.js"></script>
	<script>
		KindEditor.ready(function(K) {
			var editor = K.create('textarea[name="memo"]', {
				cssPath : '../KindEditor/plugins/code/prettify.css',
				uploadJson : '../KindEditor/asp/upload_json.asp?action=infopic',
				afterBlur: function(){this.sync();},//失去焦点执行this获得值
				allowFileManager : true,				
				afterCreate : function() {
					var self = this;
					K.ctrl(document, 13, function() {
						self.sync();
						K('form[name=example]')[0].submit();
					});
					K.ctrl(self.edit.doc, 13, function() {
						self.sync();
						K('form[name=example]')[0].submit();
					});
				}
			});
			prettyPrint();
		});
	</script>
<!--kindeditor编辑器END-->
</head>

<body topmargin="0" leftmargin="0" background="img/bg1.gif">

<div align="center">

  <center>
  <table width="760" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
    <tr>
      <td width="214" valign="top"><div align="center">
      <!--#include file=userleft.asp--></div></td>
	  <td width="5">&nbsp;</td>
      <td width="760" align="right" valign="top" bgcolor="#FFFFFF"><table width="760"  border="0" cellspacing="0" cellpadding="0" class="ty1">
        <tr>
          <td><form action="edit_info_save.asp?<%=C_ID%>" name="myform" method="POST" language="javascript" onSubmit="return formcheck()">
            <table width="100%" height="100%" border="0" align="right" cellpadding="6" cellspacing="0">
                <tr bgcolor="#FAFAFA">
                  <td height="25" colspan="2"><p align="left" style="font-weight: bold; color: #FF0000"> <span class="style5">　修改信息：</span></td>
                </tr>
                <tr bgcolor="#DBDBDB">
                  <td height="1" colspan="2"></td>
                </tr>
                <tr>
                  <td height="10" colspan="2"></td>
                </tr>
<%dim biaoti,scjiage,jiage,memo,name,email,dianhua,zt,CompPhone,youbian,dizhi,fbsj,dqsj,hfje,fbts
id=trim(request("id"))
set rs=server.createobject("adodb.recordset")
sql = "select * from [DNJY_info] where username='"&username&"' and id="&cstr(id)
rs.open sql,conn,1,1
if rs.eof then
errdick(0)
response.end
end if

biaoti=rs("biaoti")
keywords=rs("keywords")
leixing=rs("leixing")
scjiage=rs("scjiage")
jiage=rs("jiage")
memo=rs("memo")
name=rs("name")
email=rs("email")
dqsj=rs("dqsj")
fbsj=rs("fbsj")
CompPhone=rs("CompPhone")
qq=rs("qq")
youbian=rs("youbian")
dizhi=rs("dizhi")
dianhua=rs("dianhua")
hfje=rs("hfje")
fbts=rs("fbts")
zt=rs("zt")
fbsj=rs("fbsj")
%>
               <tr>
                  <td height="25">交易地区：</td>
                  <td height="25">
<%Dim rsi
set rsi=server.createobject("adodb.recordset")
set rsi=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")
%>
                                   <SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
		//dim count:
		count = 0
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
                                   <SELECT name="city_one" size="1" id="select2" onChange="changelocation(document.myform.city_one.options[document.myform.city_one.selectedIndex].value)">
                                     <OPTION value="" selected>选择城市</OPTION>
                                     <%set rsi=conn.execute("select * from DNJY_city where id>0 and twoid=0 order by indexid asc")
if rsi.eof and rsi.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsi.eof%>
                                     <OPTION value="<%=rsi("id")%>" <%if rsi("id")=rs("city_oneid") then%>selected<%end if%>><%=rsi("city")%></OPTION>
                                     <% rsi.movenext
    loop
end if
rsi.close
set rsi = nothing
%>
                                   </SELECT>
                                   <SELECT name="city_two" id="city_two" onChange="changelocation4(document.myform.city_two.options[document.myform.city_two.selectedIndex].value,document.myform.city_one.options[document.myform.city_one.selectedIndex].value)">
                                     <OPTION value="" selected>选择城市</OPTION>
                                     <%set rsi=conn.execute("select * from DNJY_city where id="&rs("city_oneid")&" and twoid<>0 and threeid=0")
if rsi.eof and rsi.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsi.eof%>
                                     <OPTION value="<%=rsi("twoid")%>" <%if rsi("twoid")=rs("city_twoid") then%>selected<%end if%>><%=rsi("city")%></OPTION>
                                     <% rsi.movenext
    loop
	end if
rsi.close
set rsi = nothing%>
                                   </SELECT>
                                   <SELECT name="city_three" id="city_three">
                                     <OPTION value="" selected>选择城市</OPTION>
                                     <%set rsi=conn.execute("select * from DNJY_city where id="&rs("city_oneid")&" and twoid="&rs("city_twoid")&" and threeid<>0")
if rsi.eof and rsi.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsi.eof%>
                                     <OPTION value="<%=rsi("threeid")%>" <%if rsi("threeid")=rs("city_threeid") then%>selected<%end if%>><%=rsi("city")%></OPTION>
                                     <% rsi.movenext
    loop
	end if
rsi.close
set rsi = nothing%>
                                   </SELECT> <font color="#ff0000"> *</font></td>
                </tr>                <tr>
                  <td width="72" height="25"> 信息类别：</td>
                  <td height="25">
<%set rsi=conn.execute("select * from DNJY_type where id>0 and twoid>0 and threeid=0 order by indexid")%>
                                 <SCRIPT language = "JavaScript">
var onecount2;
onecount2=0;
subcat2 = new Array();
        <%
		dim count2:count2 = 0
        do while not rsi.eof 
        %>
subcat2[<%=count2%>] = new Array("<%=rsi("name")%>","<%=rsi("id")%>","<%=rsi("twoid")%>");
        <%
        count2 = count2 + 1
        rsi.movenext
        loop
        rsi.close
		set rsi = nothing
        %>
onecount2=<%=count2%>;
                           </SCRIPT>
                                 <%
set rsi=conn.execute("select * from DNJY_type where id>0 and twoid>0 and threeid>0 order by indexid")
%>
                                 <SCRIPT language = "JavaScript">
var onecount3;
onecount3=0;
subcat3 = new Array();
        <%
		dim count3:count3 = 0
        do while not rsi.eof 
        %>
subcat3[<%=count3%>] = new Array("<%=rsi("name")%>","<%=rsi("id")%>","<%=rsi("twoid")%>","<%=rsi("threeid")%>");
        <%
        count3 = count3 + 1
        rsi.movenext
        loop
        rsi.close
		set rsi = nothing
        %>
onecount3=<%=count3%>;



function changelocation2(locationid)
    {
    document.myform.type_two.length = 0; 
    document.myform.type_two.options[0] = new Option('选择信息','');
	document.myform.type_three.length = 0; 
    document.myform.type_three.options[0] = new Option('选择信息','');
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
    document.myform.type_three.options[0] = new Option('选择信息','');
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
                                 <SELECT name="type_one" size="1" id="select3" onChange="changelocation2(document.myform.type_one.options[document.myform.type_one.selectedIndex].value)">
                                   <OPTION value="" selected>选择信息</OPTION>
                                   <%set rsi=conn.execute("select * from DNJY_type where id>0 and twoid=0 order by indexid asc")
if rsi.eof and rsi.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsi.eof%>
                                   <OPTION value="<%=rsi("id")%>" <%if rsi("id")=rs("type_oneid") then%>selected<%end if%>><%=rsi("name")%></OPTION>
                                   <% rsi.movenext
    loop
	end if
rsi.close
set rsi= nothing
%>
                                 </SELECT>
                                 <SELECT name="type_two" id="select4" onChange="changelocation3(document.myform.type_two.options[document.myform.type_two.selectedIndex].value,document.myform.type_one.options[document.myform.type_one.selectedIndex].value)">
                                   <OPTION value="" selected>选择信息</OPTION>
                                   <%
	set rsi=conn.execute("select * from DNJY_type where id="&rs("type_oneid")&" and twoid<>0 and threeid=0")
if rsi.eof and rsi.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsi.eof%>
                                   <OPTION value="<%=rsi("twoid")%>" <%if rsi("twoid")=rs("type_twoid") then%>selected<%end if%>><%=rsi("name")%></OPTION>
                                   <% rsi.movenext
    loop
	end if
rsi.close
set rsi= nothing%>
                                 </SELECT>
                                 <SELECT name="type_three" id="select5">
                                   <OPTION value="" selected>选择信息</OPTION>
                                   <%set rsi=conn.execute("select * from DNJY_type where id="&rs("type_oneid")&" and twoid="&rs("type_twoid")&" and threeid<>0")
if rsi.eof and rsi.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsi.eof%>
                                   <OPTION value="<%=rsi("threeid")%>" <%if rsi("threeid")=rs("type_threeid") then%>selected<%end if%>><%=rsi("name")%></OPTION>
                                   <% rsi.movenext
    loop
	end if
rsi.close
set rsi= Nothing
%>
                                 </SELECT>
                      <input type="hidden" name="id" size="7" maxlength="6" value="<%=id%>"> <font color="#ff0000"> *</font></td>
                </tr>
                <tr>
                  <td height="25"> 信息标题：</td>
                  <td height="25"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" maxlength="100" name="biaoti" size="50" value="<%=biaoti%>">
                   <font color="#ff0000"> *</font>  （<font color="#CC5200">100字节以内</font>）</td>
                </tr>
                <tr>
                  <td height="25"> 关 键 词：</td>
                  <td height="25"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" maxlength="100" name="keywords" size="50" value="<%=keywords%>">
                   <font color="#ff0000"> *</font>  <input type="button" name="btn" value="从标题复制" title="复制标题到关键词" onClick="CopyWebbiaoti(document.myform.biaoti.value);">（<font color="#CC5200">100字节以内</font>）</td>
                </tr>
                <tr>
                  <td height="25"> 交易类型：</td>
                  <td height="25"><%
				  Dim leixinglx,selectedlx
set rslx=server.createobject("adodb.recordset")
sqllx = "select leixing from DNJY_config "
rslx.open sqllx,conn,1,1
leixinglx=split(rslx("leixing"),"|")
selectedlx=""
response.write "<select name=""leixing""><option value=>类型</option>"
for x=0 to ubound(leixinglx)
If leixinglx(x)=leixing Then
selectedlx="selected"
Else
selectedlx=""
End if
response.write "<option value="""&leixinglx(x)&""" "&selectedlx&">"&leixinglx(x)&"</option>"
next
response.write "</select>"
rslx.close
set rslx=nothing%> <font color="#ff0000"> *</font></td>
                </tr>
               <!-- <tr>
                  <td height="25"> 市场价格：</td>
                  <td height="25"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="scjiage" size="7" maxlength="6" value="<%=scjiage%>" onkeyup="value=value.replace(/[^\d|^\.]/g,'')"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))"> <font color="#ff0000"> *</font> 
                    元&nbsp;&nbsp;&nbsp; （0<font color="#CC5200">元表示面议</font>）</td>
                </tr>-->
                 <tr>
                  <td height="25"> 交易价格：</td>
                  <td height="25"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="jiage" size="7" maxlength="6" value="<%=jiage%>" onkeyup="value=value.replace(/[^\d|^\.]/g,'')"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))"> <font color="#ff0000"> *</font> 
                    元&nbsp;&nbsp;&nbsp; （0<font color="#CC5200">元表示面议</font>）</td>
                </tr>
 
                <tr>
                  <td height="25">交易状态： </td>
                  <td height="25"><select name="zt" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;">
                      <option <%if zt=0 then%>selected<%end if%> value=0>还未完成交易</option>
                      <option <%if zt=1 then%>selected<%end if%> value=1>已经完成交易</option>
                    </select>
                    （<font color="#CC5200">如果此商品已经销售或者购买成功，请根据实际情况修改</font>） </td>
                </tr>
<%sj=DateDiff("d",""&fbsj&"",""&dqsj&"")%>
                      <tr>
                        <td height="25">发布天数：</td>
                        <td width="371" height="25"><select name="days" style="BACKGROUND-COLOR: #FFCCFF; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;"> 
						    <option value="1" <% if sj>=1 Then Response.write("Selected") %>>一天</option>
						    <option value="2" <% if sj>=2 Then Response.write("Selected") %>>二天</option>
						    <option value="3" <% if sj>=3 Then Response.write("Selected") %>>三天</option>
						    <option value="4" <% if sj>=4 Then Response.write("Selected") %>>四天</option>
						    <option value="5" <% if sj>=5 Then Response.write("Selected") %>>五天</option>
						    <option value="6" <% if sj>=6 Then Response.write("Selected") %>>六天</option>
                            <option value="7" <% if sj>=7 Then Response.write("Selected") %>>一星期</option>
							<option value="14" <% if sj>=14 Then Response.write("Selected") %>>二星期</option>
							<option value="21" <% if sj>=21 Then Response.write("Selected") %>>三星期</option>                            
                            <option value="30" <% if sj>=30 Then Response.write("Selected") %>>一个月</option>
							<option value="60" <% if sj>=60 Then Response.write("Selected") %>>二个月</option>
                            <option value="90" <% if sj>=90  Then Response.write("Selected") %>>三个月</option>
							<option value="180" <% if sj>=180  Then Response.write("Selected") %>>半年</option>
                            <option value="365" <% if sj=365 Then Response.write("Selected") %>>一年</option>
							<option value="1100" <% if sj=1100 Then Response.write("Selected") %>>三年</option>
                          </select>
                           </td>
                      </tr>
					                       <tr>
<td height="25" colspan="2">===============<b><font color="#ff0000">竞价信息有偿服务选项，若竞价金额大于0元系统将自动从您的帐户中等额扣款</font></b>==============<br>竞价信息可在醒目位置显示。每天均价（竞价金额÷发布天数）越高，显示越靠前，请根据发布天数选择适当竞价金额<br>
此信息竞价余额<%=FormatNumber((dqsj-now())*Round(hfje/fbts,2),2,-1)%>元，若您在下方的“竞价金额”中增加，将与余额相加并重新计算均价</td>
 </tr>
                     <tr>
                        <td height="20">竞价金额：</td>
                        <td width="450" height="-1"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="hfje" size="23" value="0" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"> 元 (您的个人帐户余款为<%=Amount%>元)</td>
                      </tr> 

				<tr>
<td height="20" colspan="2">=======================================<b><font color="#ff0000">竞价信息有偿服务选项结束</font></b>======================================</td>
 </tr>
 					  					    <tr>
                        <td height="25">阅读权限：</td>
                        <td width="371" height="25"><input type="radio" name="Readinfo" value="0" <%if rs("Readinfo")=0 then%>checked<%end if%>>游客	<input type="radio" name="Readinfo" value="1" <%if rs("Readinfo")=1 then%>checked<%end if%>>普通会员	<input type="radio" name="Readinfo" value="2" <%if rs("Readinfo")=2 then%>checked<%end if%>>VIP会员	
                             <br>（限制信息详细说明及联络方式，会员级别高向低兼容）</td>
                      </tr>
                      <!--<tr> 
                        <td height="25">
                        阅读权限：</td>
                        <td height="0" ><input type="radio" name="Readinfo" value="1" <%if rs("Readinfo")=1 then%>checked<%end if%>>普通会员	<input type="radio" name="Readinfo" value="2" <%if rs("Readinfo")=2 then%>checked<%end if%>>VIP会员	
                             <br>（限制信息详细说明及联络方式，会员级别高向低兼容）</td>
                      </tr>-->	
                <tr>
                  <td height="25"><p>详细说明：<br><%If vip=0 Then%>(信息显示时<br>屏蔽Html代码)<%End if%></p></td>
                  <td height="25"><textarea id="editor" name="memo" style="width:670px;height:400px;"><%=HTMLDecode(memo)%></textarea><input type="checkbox" name="T_SaveImg" value="1" /> 保存内容中远程图片到本地&nbsp;&nbsp;<input type="checkbox" name="AspJpeg" value="1" checked/>给上传图片加水印&nbsp;&nbsp;<input type="checkbox" name="T_conpic" value="1" />提取内容第一张图片为主图（若已上传主图此项无效）<font color="#ff0000"> *</font></td>
                </tr>
                 <!--<tr>
                  <td height="25">外观颜色：</td>
                  <td height="25"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="wcolor" size="23" maxlength="15" value="<%=rs("wcolor")%>"> &lt;=【<FONT color=blue><FONT 
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
                  <td height="25"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="ncolor" size="23" maxlength="15" value="<%=rs("ncolor")%>"> &lt;=【<FONT color=blue><FONT 
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
                <tr>
                  <td height="25">真实姓名：</td>
                  <td height="25"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="name" size="23" maxlength="15" value="<%=name%>"> <font color="#ff0000"> *</font></td>
                </tr>
                       <tr>
                        <td height="25">移动电话：</td>
                        <td height="25"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="CompPhone" size="23" maxlength="40" value="<%=CompPhone%>" onkeyup="value=value.replace(/[^\d|^\.]/g,'')"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))"> <font color="#0000ff"> *</font> （<font color="#CC5200">固定电话和移动电话必填其一</font>）</td>
                      </tr>
                <tr>
                  <td height="25"> 固定电话：</td>
                  <td height="25"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="dianhua" size="23" maxlength="40" value="<%=dianhua%>">  <font color="#0000ff"> *</font> （<font color="#CC5200">固定电话和移动电话必填其一</font>）<br>(国家代号-区号-电话-分机格式(可只填区号+电话或电话)：+086-010-85858585-2538)</td>
                </tr>
                <tr>
                  <td height="25"> 传真号码：</td>
                  <td height="25"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="fax" size="23" maxlength="40" value="<%=rs("fax")%>">  (国家代号-区号-电话-分机格式(可只填区号+电话或电话)：+086-010-85858585-2538)</td>
                </tr>
                <tr>
                  <td height="25">电子信箱：</td>
                  <td height="25"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="email" size="23" maxlength="40" value="<%=email%>">
                    <font color="#ff0000"> *</font> （<font color="#CC5200">该邮件地址将接收你的交易信息，请正确填写！</font>）</td>
                </tr>					  
                      <tr>
                        <td height="25">QQ 号 码：</td>
                        <td width="371" height="-1"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="qq" size="23" maxlength="40" value="<%=qq%>" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"></td>
                      </tr>
                <tr>
                  <td height="25"> 微信号码：</td>
                  <td height="25"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="weixin" size="23" maxlength="50" value="<%=rs("weixin")%>"></td>
                </tr>					  
                      <tr>
                        <td height="25">邮政编码：</td>
                        <td width="371" height="-1"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="youbian" size="23" maxlength="6" value="<%=youbian%>" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"></td>
                      </tr>
                      <tr>
                        <td height="25" width="71"> 公司名称：</td>
                        <td height="25" width="371">
                            <input name="mpname" type="text" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" size="50" maxlength="50" value="<%=rs("mpname")%>">
                          &nbsp; 
                          </td>
                      </tr>					  
                      <tr>
                        <td height="25" width="71"> 联系地址：</td>
                        <td height="25" width="371">
                            <input name="dizhi" type="text" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" size="50" maxlength="30" value="<%=dizhi%>">
                          &nbsp; 
                         </td>
                      </tr>

                <tr>
                  <td height="4" colspan="2"></td>
                </tr>
                <tr>
                  <td height="25" colspan="2" align="center"><table width="50%"  border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><div align="center">
                            <input name="I2" type="submit" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #a2a2a2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #a2a2a2 1px solid; BORDER-TOP: #ffffff 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" onClick="javascript:return formcheck();" value="确认编辑" border="0">&nbsp;&nbsp;&nbsp;<%If gxkf>0 and DateDiff("d",fbsj,Now())>1 then%><font color="#ff0000">编辑更新本条信息将扣除积分<%=gxkf%>分 <input type="hidden" name="gxkfto" value="<%=gxkf%>"><%end if%><%If (vip=0 and vipno=0) or onoff=0 then%><font color="#ff0000">修改后将重新审核</font><%end if%>
                        </div></td>
                        <td><div align="center">
                            <input name="I2" type="reset" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #a2a2a2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #a2a2a2 1px solid; BORDER-TOP: #ffffff 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" onClick="javascript:return CheckForm();" value="还原编辑" border="0">
                        </div></td>
                      </tr>
                  </table></td>
                </tr>
              </table>
          </form></td>
        </tr>
      </table></td>
      <td width="4" align="right" valign="top" bgcolor="#e6e6e6"></td>
    </tr>
    <tr><%Call CloseRs(rs)
%>
      <td height="26" colspan="3"><!--#include file=footer.asp--></td>
    </tr>
  </table>
  </center>
</div>
</body>
</html>

