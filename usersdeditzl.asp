<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=top.asp-->
<!--#include file=usercookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>

<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">

<style type="text/css">
<!--
.style5 {color: #FF0000; font-weight: bold; }
.style6 {color: #FF0000}
-->
</style>
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
function textLimitCheck2(thisArea, maxLength){
    if (thisArea.value.length > maxLength){
      alert(maxLength + ' 个字限制. \r超出的将自动去除.');
      thisArea.value = thisArea.value.substring(0, maxLength);
      thisArea.focus();
    }
    /*回写span的值，当前填写文字的数量*/    
	messageCount2.innerText = thisArea.value.length;
  }
</script>
</head>

<body background="img/bg1.gif" leftmargin="0" topmargin="0">

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="980" height="426">
    <tr>
      <td width="214" height="354" valign="top" bgcolor="#FFFFFF"><div align="center">
      
        <!--#include file=userleft.asp--></div></td>
		<td width="5">&nbsp;</td>
      <td width="760" height="354" valign="top" bgcolor="#FFFFFF"><table width="99%" height="100%" border="0" align="right" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse" class="ty1">
        <tr>
          <td width="100%" height="362" align="center" valign="top"><div align="right">
              <table width="100%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse">
                <form method="POST" name="thisForm" action="usersdeditzlchk.asp?<%=C_ID%>" onSubmit="return CheckForm();">
                  <tr>
                    <td height="25" bgcolor="#FAFAFA"><span class="style5">申请或修改店企注意事项：</span></td>
                    <td bgcolor="#FAFAFA">&nbsp;</td>
                  </tr>
                  <tr>
                    <td height="25" >一、网上店企是利用本站平台搭建相对独立的网络店铺或发布企业名片，您可通过自己的网络店铺店或企业名片发布信息和管理、经营商品。<br>升级为VIP会员将得到更多的权限。<a href="upgradevip.asp?<%=C_ID%>"><font color=#0000ff>我要升级为VIP</font></a><br>二、申请要上传相关证照才能通过审核。若您未曾上传过，请<a href="certificate.asp?<%=C_ID%>"><font color=#0000ff>上传</font></a><br>三、申请（修改）后须审核才生效。</td>
                    <td bgcolor="#FAFAFA">&nbsp;</td>
                  </tr> 
                  </tr>
                  <tr>
                    <td height="25" ><b>审核结果通知：</b><%=rs("notification")%></td>
                    <td bgcolor="#FAFAFA">&nbsp;</td>
                  </tr>   				  
                  <tr bgcolor="#CCCCCC">
                    <td width="760" height="1">
 <%
dim m,sdkg,sdfg,mpkg
username=request.cookies("dick")("username")
set rs=conn.execute("select count(id) from [DNJY_info] where yz=1 and username='"&username&"'")
m=rs(0)
rs.close
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_user] where username='"&username&"'" 
rs.open sql,conn,1,1
if rs.eof then
response.write "<br>"
response.write "<li>参数错误！"
response.end
end if
sdkg=rs("sdkg")
sdfg=rs("sdfg")
mpkg=rs("mpkg")
mpfg=rs("mpfg")
%>
<SCRIPT language=javascript>
<!--
//字节数限制，与下面配合=========
String.prototype.Length=function(){
temp=this.replace(/([^\x00-\xff])/g,"$1$1");
return temp.length
}
//显示输入字符数

function kn()
  {
   var ln=thisForm.sdname.value.Length();
     num.innerText='您还可以输入:'+(30-ln)+'个字符';
      //if (ln>=8) form.sdname.readOnly=true;  //这行是如果满足条件表单无法输入和修改
}

function kn2()
  {
   var ln=thisForm.mpname.value.Length();
     num2.innerText='您还可以输入:'+(30-ln)+'个字符';
      //if (ln>=8) form.sdname.readOnly=true;  //这行是如果满足条件表单无法输入和修改
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



function CheckForm()
{
if (document.thisForm.city_one.value==0)
	{
	  alert("请选择地区！")
	  document.thisForm.city_one.focus()
	  return false
	 }
if (document.thisForm.type_one.value==0)
	{
	  alert("请选择行业分类！")
	  document.thisForm.type_one.focus()
	  return false
	 }
if ((document.thisForm.sdname.value=="") && (document.thisForm.sdkg.value=="2" )) 
     {
	 alert("店名必须填写！");
	  document.thisForm.sdname.focus()
	  return false
  }	
if (document.thisForm.sdname.value.Length()>30)  //字节数限制，与上面配合
     {
	 alert("店名长度要求30字节内，请重新输入！");
	  document.thisForm.sdname.focus()
	  return false
  }	 
if ((document.thisForm.sdname.value.length>1) && (document.thisForm.sdfg.value=="0" ))
	{
	    alert("请选择网店风格");
		document.thisForm.sdfg.focus();
	    return false;
	}
if ((document.thisForm.sdname.value.length>1) && (document.thisForm.sdjj.value=="" ))
	{
	    alert("网店简介没有填写!");
		document.thisForm.sdjj.focus();
	    return false;
	}
if (document.thisForm.mpname.value=="" && (document.thisForm.mpkg.value=="2" )) 
     {
	 alert("名片名称必须填写！");
	  document.thisForm.mpname.focus()
	  return false
  }	
if (document.thisForm.mpname.value.Length()>30)  //字节数限制，与上面配合
     {
	 alert("名片名称长度要求30字节内，请重新输入！");
	  document.thisForm.mpname.focus()
	  return false
  }
  
if ((document.thisForm.mpname.value.length>1) && (document.thisForm.mpfg.value=="0" ))
	{
	    alert("请选择名片风格");
		document.thisForm.mpfg.focus();
	    return false;
	}
if ((document.thisForm.mpname.value.length>1) && (document.thisForm.mpjj.value=="" ))
	{
	    alert("名片简介没有填写!");
		document.thisForm.mpjj.focus();
	    return false;
	}  
if(document.thisForm.name.value.length<1 )
	{
	    alert("请填联系人!");
		document.thisForm.name.focus();
	    return false;
	}

if(document.thisForm.dianhua.value.length>30 )
	{
	    alert("固定电话不能超过30个字符!");
		document.thisForm.dianhua.focus();
	    return false;
	}
if(document.thisForm.dianhua.value=="" && document.thisForm.CompPhone.value=="") 
	{
	    alert("固定电话和移动电话不能同时为空!");
		document.thisForm.dianhua.focus()
	    return false;
	}
//var sMobile = document.thisForm.dianhua.value
//if((document.thisForm.dianhua.value!="") && (!(/^(0[0-9]{2,3}\-)?([2-9][0-9]{6,7})+(\-[0-9]{1,4})?$/.test(sMobile)))) //(区号-电话-分机格式：010-85858585-2538)
//	{
//		alert("不规范的电话号码");
//		document.thisForm.dianhua.focus();
//		return false;
//	}	
//	var sMobile = document.thisForm.CompPhone.value
//	if((document.thisForm.CompPhone.value!="") && (!(/^((\(\d{3}\))|(\d{3}\-))?13[0-9]\d{8}|15[0-9]\d{8}|18[0-9]\d{8}/.test(sMobile)))) //(检验13，15，18号段)
//	{
//		alert("不是完整的11位手机号或者正确的手机号前七位");
//		document.thisForm.CompPhone.focus();
//		return false;
//	}
//var sMobile = document.thisForm.fax.value
//if((document.thisForm.fax.value!="") && (!(/^([0\+]\d{2,3}-)?(0[0-9]{2,3}\-)?([2-9][0-9]{6,7})+(\-[0-9]{1,4})?$/.test(sMobile)))) //(国家代号-区号-电话-分机格式(可选)：+086-010-85858585-2538)
//	{
//		alert("不规范的传真电话号码");
//		document.thisForm.fax.focus();
//		return false;
//	}
if(document.thisForm.fax.value.length>30 )
	{
	    alert("传真不能超过30个字符!");
		document.thisForm.fax.focus();
	    return false;
	}
if(document.thisForm.email.value.length>30 )
	{
	    alert("信箱不能超过30个字符!");
	    document.thisForm.email.focus();
	    return false;
	}
   if((!checkemail(thisForm.email.value))&&(document.thisForm.email.value!=""))
	{
    alert("您输入Email地址不正确，请重新输入!");
    document.thisForm.email.focus();
    return false;
    }
if(document.thisForm.youbian.value.length>8 )
	{
	    alert("邮编不能超过8个字符!");
		document.thisForm.youbian.focus();
	    return false;
	}
if(document.thisForm.dizhi.value.length>100 )
	{
	    alert("地址不能超过100个字符!");
		document.thisForm.dizhi.focus();
	    return false;
	}
if(document.thisForm.http.value.length>100 )
	{
	    alert("网址不能超过100个字符!");
		document.thisForm.http.focus();
	    return false;
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
if(contain(document.thisForm.http.value,"http://"))
{ 
alert("网址前不要带http://");
document.thisForm.http.focus();
return false;
}
if(contain(document.thisForm.http.value,"/"))
{ 
alert("网址后不要带/");
document.thisForm.http.focus();
return false;
}    
if(document.thisForm.zhuying.value.length>50 )
	{
	    alert("经营范围不能超过50个字符!");
		document.thisForm.zhuying.focus();
	    return false;
	}							

}

//-->
                    </SCRIPT></td>
                    <td width="17" height="1"></td>
                  </tr>
                  <tr>
                    <td width="760" valign="top">
					<table width="760" height="105" border="0" align="left" cellpadding="5" cellspacing="0">
				
                        <tr>
                          <td height="10" colspan="3"><%if m<xinxis then
response.write "很抱歉，您发布的信息只有 "&m&" 条，要发布 "&xinxis&" 条以上有价值并经审核通过的信息才能申请店铺或发布企业名片，继续努力噢！"%>

<%else %></td>
                        </tr>
	                  <tr><td height="30" width="60" > </td>
                     <td height="30"  colspan="3">(带<font color="#ff0000">*</font>必填)</td>
                     </tr>
                        <tr>
                          <td height="30"><p align="right">所属地区：</td>
                          <td height="30" width="700"><%Dim rsi
set rsi=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0 order by indexid")
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
set rsi=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid>0 order by indexid")
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
    document.thisForm.city_two.length = 0; 
	document.thisForm.city_two.options[0] = new Option('选择城市','');
	document.thisForm.city_three.length = 0; 
	document.thisForm.city_three.options[0] = new Option('选择城市','');
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
    document.thisForm.city_three.options[0] = new Option('选择城市','');
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
      <SELECT name="city_one" size="1" class="inputa" id="select2" onChange="changelocation(document.thisForm.city_one.options[document.thisForm.city_one.selectedIndex].value)">
            <OPTION value="" selected>选择城市</OPTION>
  <%set rsi=conn.execute("select * from DNJY_city where id>0 and twoid=0 and threeid=0 order by indexid")
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
	  <SELECT name="city_two" id="city_two" class="inputa" onChange="changelocation4(document.thisForm.city_two.options[document.thisForm.city_two.selectedIndex].value,document.thisForm.city_one.options[document.thisForm.city_one.selectedIndex].value)">
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
	     <SELECT name="city_three" id="city_three" class="inputa">
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
    </SELECT>
                           <font color="#ff0000">*</font></td>
                        </tr>
<tr>
                  <td height="25" align="right"> 行业类别：</td>
                  <td height="25">
<%set rsi=conn.execute("select * from DNJY_hy_type where id>0 and twoid>0 and threeid=0 order by indexid")%>
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
set rsi=conn.execute("select * from DNJY_hy_type where id>0 and twoid>0 and threeid>0 order by indexid")
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
    document.thisForm.type_two.length = 0; 
    document.thisForm.type_two.options[0] = new Option('选择行业二级分类','');
	document.thisForm.type_three.length = 0; 
    document.thisForm.type_three.options[0] = new Option('选择行业三级分类','');
    var locationid=locationid;
	var i;
    for (i=0;i < onecount2; i++)
        {
            if (subcat2[i][1] == locationid)
            { 
                document.thisForm.type_two.options[document.thisForm.type_two.length] = new Option(subcat2[i][0], subcat2[i][2]);
            }        
        }
       
    }    
     function changelocation3(locationid,locationid1)
    {
    document.thisForm.type_three.length = 0; 
    document.thisForm.type_three.options[0] = new Option('选择行业三级分类','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount3; i++)
        {
            if (subcat3[i][2] == locationid)
            { 
			if (subcat3[i][1] == locationid1)
			{
                document.thisForm.type_three.options[document.thisForm.type_three.length] = new Option(subcat3[i][0], subcat3[i][3]);
            }        
        }
       } 
    }            
	                       </SCRIPT>
                                 <SELECT name="type_one" size="1" id="select3" class="inputa" onChange="changelocation2(document.thisForm.type_one.options[document.thisForm.type_one.selectedIndex].value)">
                                   <OPTION value="" selected>选择行业一级分类</OPTION>
                                   <%set rsi=conn.execute("select * from DNJY_hy_type where id>0 and twoid=0 order by indexid asc")
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
                                 <SELECT name="type_two" id="select4" class="inputa" onChange="changelocation3(document.thisForm.type_two.options[document.thisForm.type_two.selectedIndex].value,document.thisForm.type_one.options[document.thisForm.type_one.selectedIndex].value)">
                                   <OPTION value="" selected>选择行业二级分类</OPTION>
                                   <%
	set rsi=conn.execute("select * from DNJY_hy_type where id="&rs("type_oneid")&" and twoid<>0 and threeid=0")
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
                                 <SELECT name="type_three" id="select5"  class="inputa">
                                   <OPTION value="" selected>选择行业三级分类</OPTION>
                                   <%set rsi=conn.execute("select * from DNJY_hy_type where id="&rs("type_oneid")&" and twoid="&rs("type_twoid")&" and threeid<>0")
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
                      <font color="#ff0000">*</font></td>
                </tr>
                        <tr bgcolor="#BBCEFD">
                          <td height="30"><p align="right">网店名称：</td>
                          <td height="30" width="700"><input name="sdname" type="text" class="inputa" value="<%=rs("sdname")%>" size="30" maxlength="30" onpropertychange="kn()"> <font color="#ff0000">*</font>&nbsp;&nbsp;<span id=num>你还可以输入30个字符（15个汉字）</span>
                           </td>
                        </tr>
                        <tr bgcolor="#BBCEFD">
                          <td height="30" align="right">网店风格：</td>
                          <td height="30"><select name="sdfg"  class="inputa">
						    <option  value="0" style="background-color:#F8F4F5 ;color: #FF0000">选择网店风格</option>
                            <option  value="1" <%if sdfg=1 then%>selected<%end if%>>蓝天白云</option>
                            <option  value="2" <%if sdfg=2 then%>selected<%end if%>>黄金季节</option>
							<option  value="3" <%if sdfg=3 then%>selected<%end if%>>绿草如茵</option>
							<option  value="4" <%if sdfg=4 then%>selected<%end if%>>万紫千红</option>
                          </select>
&nbsp;&nbsp; <a href="img/dqmb.jpg" target="_blank">效果预览</a></td>
                        </tr>
                         <tr bgcolor="#BBCEFD">
                          <td height="30"><p align="right">网店状态：</td>
                          <td height="30" width="384"><label>
                            <select name="sdkg" class="inputa">
                              <option value="0" <%if sdkg=0 then%>selected<%end if%>>关闭状态</option>
                              <option value="2" <%if sdkg=2 then%>selected<%end if%>>申请开放</option>
                                                        </select>
                          当前状态：
<%if sdkg=1 then
response.write "<font color=""#ff0000""><strong>开放状态</strong></font>"
elseif sdkg=0 then
response.write "<font color=""#0066FF""><strong>关闭状态</strong></font>"
elseif sdkg=2 then
response.write "<font color=""#FF9900""><strong>申请开放</strong></font>"
end if%>
                          </label></td>
                        </tr>
                        <tr bgcolor="#BBCEFD">
                          <td width="120" height="30" valign="top"><p align="right">
                            <p align="right">网店简介：</td>
                          <td width="300" height="30" valign="top"><textarea name="sdjj" cols="65" rows="10" onkeyUp="textLimitCheck(this, 800);"><%=rs("sdjj")%></textarea> 限 800 个字符 已输入 <font color="#CC0000"><span id="messageCount">0</span></font> 个字</td>
                        </tr>
                        <tr bgcolor="#E6F0FA">
                          <td height="30"><p align="right">名片名称：</td>
                          <td height="30" width="700"><input name="mpname" type="text" class="inputa" value="<%=rs("mpname")%>" size="30" maxlength="30" onpropertychange="kn2()"> <font color="#ff0000">*</font>&nbsp;&nbsp;<span id=num2>你还可以输入30个字符（15个汉字）</span>
                           </td>
                        </tr>
                        <tr bgcolor="#E6F0FA">
                          <td height="30" align="right">名片风格：</td>
                          <td height="30"><select name="mpfg"  class="inputa">
						    <option  value="0" style="background-color:#F8F4F5 ;color: #FF0000">选择企业名片风格</option>
                            <option  value="1" <%if mpfg=1 then%>selected<%end if%>>蓝天白云</option>
                            <option  value="2" <%if mpfg=2 then%>selected<%end if%>>深蓝宝石</option>
							<option  value="3" <%if mpfg=3 then%>selected<%end if%>>青青河草</option>
							<option  value="4" <%if mpfg=4 then%>selected<%end if%>>黄金季节</option>
							<option  value="5" <%if mpfg=5 then%>selected<%end if%>>温馨粉红</option>
                          </select>
&nbsp;&nbsp; <a href="img/dqmb.jpg" target="_blank">效果预览</a></td>
                        </tr>
                         <tr bgcolor="#E6F0FA">
                          <td height="30"><p align="right">名片状态：</td>
                          <td height="30" width="384"><label>
                            <select name="mpkg" class="inputa">
                              <option value="0" <%if mpkg=0 then%>selected<%end if%>>关闭状态</option>
                              <option value="2" <%if mpkg=2 then%>selected<%end if%>>申请开放</option>
                                                        </select>
                          当前状态：
                          <%
if mpkg=1 then
response.write "<font color=""#ff0000""><strong>开放状态</strong></font>"
elseif mpkg=0 then
response.write "<font color=""#0066FF""><strong>关闭状态</strong></font>"
elseif mpkg=2 then
response.write "<font color=""#FF9900""><strong>申请开放</strong></font>"
end if%>
                          </label></td>
                        </tr>
                        <tr bgcolor="#E6F0FA">
                          <td height="30" valign="top"><p align="right">
                            <p align="right">名片简介：</td>
                          <td width="300" height="30" valign="top"><textarea name="mpjj" cols="65" rows="10" onkeyUp="textLimitCheck2(this, 800);"><%=rs("mpjj")%></textarea> 限 800 个字符 已输入 <font color="#CC0000"><span id="messageCount2">0</span></font> 个字</td>
                        </tr>
                        <tr>
                          <td height="30" valign="top" align="right">联 系 人：</td>
                          <td height="30" valign="top"><input name="name" type="text" class="inputa"value="<%=rs("name")%>" size="20" maxlength="8"> <font color="#ff0000">*</font></td>
                        </tr>

                        <tr>
                          <td height="30" valign="top" align="right">固定电话：</td>
                          <td height="30" valign="top"><input name="dianhua" type="text" class="inputa" value="<%=rs("dianhua")%>" size="20" maxlength="30">
                             <font color="#ff0000">*</font> <font color="#CC5200">(区号-电话-分机格式：010-85858585-2538)</font> </td>
                        </tr>
                        <tr>
                          <td height="30" valign="top" align="right">移动电话：</td>
                          <td height="30" valign="top"><input name="CompPhone" type="text" class="inputa" value="<%=rs("CompPhone")%>" size="20" maxlength="20" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"> <font color="#CC5200">（固定电话和移动电话必填其一）</font></td>
                        </tr>
                        <tr>
                          <td height="30" valign="top" align="right">传真号码：</td>
                          <td height="30" valign="top"><input name="fax" type="text" class="inputa" value="<%=rs("fax")%>" size="20" maxlength="30">
                            同上<span class="style6">↑</span></td>
                        </tr>
                        <tr>
                          <td height="30" valign="top" align="right">电子邮件：</td>
                          <td height="30" valign="top"><input name="email" type="text" class="inputa"value="<%=rs("email")%>" size="30" maxlength="30"> <font color="#CC5200">（接收重要信息，请正确填写！）</font> </td>
                        </tr>
                        <tr>
                          <td height="30" valign="top" align="right">联络QQ号：</td>
                          <td height="30" valign="top"><input name="qq" type="text" class="inputa" value="<%=rs("qq")%>" size="20" maxlength="15" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')">
                            同上<span class="style6">↑</span></td>
                        </tr>                        
                        <tr>
                          <td height="30" valign="top" align="right">邮政编码：</td>
                          <td height="30" valign="top"><input name="youbian" type="text" class="inputa" value="<%=rs("youbian")%>" size="10" maxlength="8" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"></td>
                        </tr>
                        <tr>
                          <td height="30" valign="top" align="right">详细地址：</td>
                          <td height="30" valign="top"><input name="dizhi" type="text" class="inputa" value="<%=rs("dizhi")%>" size="50" maxlength="100"></td>
                        </tr>
                        <tr>
                          <td height="30" valign="top" align="right">店企网站：</td>
                          <td height="30" valign="top"><input name="http" type="text" class="inputa" value="<%=rs("http")%>" size="50" maxlength="100">
                         <br>如果您有网站可填网址，如 <span class="style6">www.ip126.com</span> 前面不能带http:// </td>
                        </tr>
                        <tr>
                          <td height="30" valign="top" align="right">主要经营：</td>
                          <td height="30" valign="top"><input name="zhuying" type="text" class="inputa" value="<%=rs("zhuying")%>" size="50" maxlength="50"></td>
                        </tr>                        
       

                        <tr>
                          <td colspan="3"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                              
                              <tr>
                                <td><div align="center">
                                    <input  name="I4" type="submit" class="inputb" value="提交修改" border="0" >
                               <span class="style6">申请（修改）后须审核才生效</span> </div> </td>
                                <td><div align="center">
                                    <input  name="I5" type="reset" class="inputb" id="I5" value="取消修改" border="0" >
                                </div></td>
                              </tr>
                          </table></td>
                        </tr>
   <%end if %>                     
                    </table></td>
                    <td width="17"> 　</td>
                  </tr>
                </form>
              </table>              
<%Call CloseRs(rs)
%>
          </div></td>
        </tr>
      </table></td>
      <td width="4" valign="top" bgcolor="#e6e6e6"></td>
    </tr>

    <tr>
      <td height="26" colspan="3"><!--#include file=footer.asp--></td>
    </tr>
  </table>
  </center>
</div>
</body>
</html>
