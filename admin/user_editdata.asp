<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Call OpenConn
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_user] where username='"&trim(request("username"))&"'" 
rs.open sql,conn,1,1
if rs.eof then
response.write "<font color=#0000FF>没有此会员帐号！</font>"
Else%>
<SCRIPT language=javascript>
<!--
//验证身份证正确性
function checkeNO(NO){
var str=NO;
//在JavaScript中，正则表达式只能使用"/"开头和结束，不能使用双引号
var Expression=/\d{17}[\d|X]|\d{15}/; 
var objExp=new RegExp(Expression);
if(objExp.test(str)==true){
return true;
}else{
return false;
}
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
    //if (document.thisForm.city_one.value==0)
	//{
	//  alert("请选择地区！")
	//  document.thisForm.city_one.focus()
	//  return false
	// }
    //if (document.thisForm.type_one.value==0)
	//{
	 // alert("请选择信息分类！")
	 // document.thisForm.type_one.focus()
	 // return false
	 //}
    //if(document.thisForm.name.value.length<1)
	//{
	//    alert("真实姓名没有填写!");
	//	document.thisForm.name.focus()
	//    return false;
	//}
	if(((document.thisForm.password.value.length<5) || (document.thisForm.password.value.length>12)) &&(document.thisForm.password.value!=""))
        {
            alert("密码长度限5－12字节!");
			document.thisForm.password.focus()
            return false;
        }
    if((!checkeNO(thisForm.idcard.value)) && (document.thisForm.idcard.value!="")){
    alert("您输入身份证号码不正确!");
    document.thisForm.idcard.focus();
    return false;
    }
	if(document.thisForm.dianhua.value=="" && document.thisForm.CompPhone.value=="") 
	{
	    alert("固定电话和手机不能同时为空!");
		document.thisForm.dianhua.focus()
	    return false;
	}
// var sMobile = document.thisForm.dianhua.value
//if((document.thisForm.dianhua.value!="") && (!(/^([0\+]\d{2,3}-)?(0[0-9]{2,3}\-)?([2-9][0-9]{6,7})+(\-[0-9]{1,4})?$/.test(sMobile)))) //(国家代号-区号-电话-分机格式(可选)：+086-010-85858585-2538)
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
	if((!checkemail(thisForm.email.value))&&(document.thisForm.email.value!=""))
    {
    alert("您输入Email地址不正确，请重新输入!");
    document.thisForm.email.focus();
    return false;
    }
}

//-->
</SCRIPT>
<meta http-equiv="Content-Language" content="zh-cn">
<title>修改资料</title>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<body topmargin="0" leftmargin="0">
<div align="center">
  <center>
            <table width="500" border="0" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
              <form method="POST" name="thisForm" action="user_editdata_save.asp?username=<%=trim(request("username"))%>">
              <tr> 
                <td width="16" background="../images/obj_waku3_03.gif">
                　</td>
                <td width="500" bgcolor="#EEEEEE">
                  <table width="500" border="0" cellspacing="0" cellpadding="0" height="297">
                    <tr>
                      <td height="22" bgcolor="#EEEEEE" width="489" colspan="6">
                      <p align="center">修改用户<%=rs("username")%>的资料</td>
                      </tr>
                       <tr>
                          <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      <p align="right">所属地区：</td>
                      <td height="30" bgcolor="#EEEEEE" width="430" colspan="5"><%Dim rsi
set rsi=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0 order by indexid")
%>
      <SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
		dim count:
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
                           （默认地区）</td>
                        <tr><td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      <p align="right">信息分类：</td>
                          <td height="30" bgcolor="#EEEEEE" width="410" colspan="5"><%Dim rsj
set rsj=conn.execute("select * from DNJY_type where id>0 and twoid>0 and threeid=0 order by indexid")
%>
      <SCRIPT language = "JavaScript">
var onecount5;
onecount5=0;
subcat5 = new Array();
        <%
		dim count5:
		count5 = 0
        do while not rsj.eof 
        %>
subcat5[<%=count5%>] = new Array("<%=rsj("name")%>","<%=rsj("id")%>","<%=rsj("twoid")%>");
        <%
        count5 = count5 + 1
        rsj.movenext
        loop
        rsj.close
		set rsj=nothing
        %>
onecount5=<%=count5%>;
</SCRIPT>
<%
set rsj=conn.execute("select * from DNJY_type where id>0 and twoid>0 and threeid>0 order by indexid")
%>
 <SCRIPT language = "JavaScript">
var onecount6;
onecount6=0;
subcat6 = new Array();
        <%
		dim count6:count6 = 0
        do while not rsj.eof 
        %>
subcat6[<%=count6%>] = new Array("<%=rsj("name")%>","<%=rsj("id")%>","<%=rsj("twoid")%>","<%=rsj("threeid")%>");
        <%
        count6 = count6 + 1
        rsj.movenext
        loop
        rsj.close
		set rsj = nothing
        %>
onecount6=<%=count6%>;

function changelocation5(locationid)
    {
    document.thisForm.type_two.length = 0; 
	document.thisForm.type_two.options[0] = new Option('选择信息分类','');
	document.thisForm.type_three.length = 0; 
	document.thisForm.type_three.options[0] = new Option('选择信息分类','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount5; i++)
        {
            if (subcat5[i][1] == locationid)
            { 
                document.thisForm.type_two.options[document.thisForm.type_two.length] = new Option(subcat5[i][0], subcat5[i][2]);
            }        
        }
        
    }    
	
	function changelocation6(locationid,locationid1)
    {
    document.thisForm.type_three.length = 0; 
    document.thisForm.type_three.options[0] = new Option('选择信息分类','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount6; i++)
        {
            if (subcat6[i][2] == locationid)
            { 
			if (subcat6[i][1] == locationid1)
			{
                document.thisForm.type_three.options[document.thisForm.type_three.length] = new Option(subcat6[i][0], subcat6[i][3]);
            }        
        }
       } 
    }            
              </SCRIPT>
      <SELECT name="type_one" size="1" class="inputa" id="select2" onChange="changelocation5(document.thisForm.type_one.options[document.thisForm.type_one.selectedIndex].value)">
            <OPTION value="" selected>选择信息分类</OPTION>
  <%set rsj=conn.execute("select * from DNJY_type where id>0 and twoid=0 and threeid=0 order by indexid")
if rsj.eof or rsj.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsj.eof%>
  <OPTION value="<%=rsj("id")%>" <%if rsj("id")=rs("type_oneid") then%>selected<%end if%>><%=rsj("name")%></OPTION>
 <%rsj.movenext
    loop
	%>
	<%end if
rsj.close
set rsj = nothing
%>
      </SELECT> 
	  <SELECT name="type_two" id="type_two" class="inputa" onChange="changelocation6(document.thisForm.type_two.options[document.thisForm.type_two.selectedIndex].value,document.thisForm.type_one.options[document.thisForm.type_one.selectedIndex].value)">
    <OPTION value="" selected>选择信息分类</OPTION>
   <%
set rsj=conn.execute("select * from DNJY_type where id="&rs("type_oneid")&" and twoid>0 and threeid=0")
if rsj.eof and rsj.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsj.eof%>
  <OPTION value="<%=rsj("twoid")%>" <%if rsj("twoid")=rs("type_twoid") then%>selected<%end if%>><%=rsj("name")%></OPTION>
 <%rsj.movenext
    loop
	%>
	<%end if
rsj.close
set rsj = nothing
%>
	</SELECT>
	     <SELECT name="type_three" id="type_three" class="inputa">
         <OPTION value="" selected>选择信息分类</OPTION>
		<%
set rsj=conn.execute("select * from DNJY_type where id="&rs("type_oneid")&" and twoid="&rs("type_twoid")&" and threeid<>0")
if rsj.eof and rsj.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsj.eof%>
<OPTION value="<%=rsj("threeid")%>" <%if rsj("threeid")=rs("type_threeid") then%>selected<%end if%>><%=rsj("name")%></OPTION>
   <% rsj.movenext
    loop
	%>
<%end if
rsj.close
set rsj = nothing%>
    </SELECT>
                           （默认分类）</td>
                        </tr>                        </tr>

                    <tr>
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      <p align="right">真实姓名：</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<input type="text" maxlength="12" name="name" size="24" value="<%=rs("name")%>"> </td>
                    </tr> 
                    <tr>
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      <p align="right">密&nbsp;&nbsp;&nbsp;&nbsp;码：</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<input type="text" name="password" size="24" value=""> 5-15位，不修改留空</td>
                    </tr>

                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      <p align="right">身 份 证：</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<input type="text" maxlength="20" name="idcard" size="24" value="<%=rs("idcard")%>" ></td>
                    </tr>
                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      <p align="right">性&nbsp;&nbsp;&nbsp;&nbsp;别：</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<select id="ctlSex" name="Sex">
                      <option <%if rs("sex")=1 then%>selected<%end if%> value="1">男</option>
                      <option <%if rs("sex")=0 then%>selected<%end if%> value="0">女</option>
                      </select></td>
                    </tr>

                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      手机号码：</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<input type="text" maxlength="30" name="CompPhone" size="24" value="<%=rs("CompPhone")%>" > <font color="#0000ff"> *</font>  手机和固定电话必填其一</td>
                    </tr>
                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      固定电话：</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<input type="text" maxlength="30" name="dianhua" size="24" value="<%=rs("dianhua")%>" > <font color="#0000ff"> *</font><br> (国家代号-区号-电话-分机格式(可选)：+086-010-85858585-2538)</td>
                    </tr>
                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      传&nbsp;&nbsp;&nbsp;&nbsp;真：</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<input type="text" maxlength="30" name="fax" size="24" value="<%=rs("fax")%>" ><br> (国家代号-区号-电话-分机格式(可选)：+086-010-85858585-2538)</td>
                    </tr>
                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      <p align="right">QQ 号码：</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<input type="text" maxlength="50" name="qq" size="24" value="<%=rs("qq")%>"></td>
                    </tr>
                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      <p align="right">微信号码：</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<input type="text" maxlength="50" name="weixin" size="24" value="<%=rs("weixin")%>"></td>
                    </tr>
                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      <p align="right">电子信箱：</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<input type="text" maxlength="50" name="email" size="24" value="<%=rs("email")%>"></td>
                    </tr>
                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      <p align="right">订阅资讯：</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<%if rs("maillist")=1 then%>               
                <input type="radio" name="maillist" value="1" id="maillist" checked>
                 订阅&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="maillist" value="0" id="maillist">
                不订</td>
                <%else%>                         
                <input type="radio" name="maillist" value="1" id="maillist">
                 订阅&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="maillist" value="0" id="maillist" checked>
                不订</td><%end if%></td>
                    </tr>                    
                    
                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      邮政编码：</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<input type="text" maxlength="6" name="youbian" size="12" value="<%=rs("youbian")%>" > </td>
                    </tr>
                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      公司名称：</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<input type="text" maxlength="100" name="mpname" size="48" value="<%=rs("mpname")%>" > 也作企业名片名称</td>
                    </tr>
                    <tr>
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right"> 通信地址：</td>
                      <td height="30" bgcolor="#EEEEEE" colspan="2">&nbsp;<input type="text" maxlength="100" name="dizhi" size="48" value="<%=rs("dizhi")%>" ></td>
                   <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      <p align="right">审核通过：</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<%if rs("useryz")=1 then%>               
                <input type="radio" name="useryz" value="1" id="useryz" checked>
                 通过&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="useryz" value="0" id="useryz">
                不通过</td>
                <%else%>                         
                <input type="radio" name="useryz" value="1" id="useryz">
                 通过&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="useryz" value="0" id="useryz" checked>
                不通过</td><%end if%></td>
                    </tr>
                    <tr> 
                      <td height="25" bgcolor="#EEEEEE" width="70">
                      <p align="center">　</td>

                      <td height="25" bgcolor="#EEEEEE" width="454">
                      <p align="center">
					  <input type="hidden" name="username" value="<%=rs("username")%>">
                      <input border="0" onClick="javascript:return CheckForm();" src="../images/Login_but.gif" name="I4" type="image"></td>
                    </tr>
                    <tr> 
                      <td height="10" bgcolor="#EEEEEE" width="454" colspan="4">　</td>
                    </tr>
                    <tr> 
                      <td height="10" bgcolor="#EEEEEE" width="454" colspan="4">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 说明：帐号不能包含下面的字符：&quot;'%#&amp;&lt;&gt;、中文以及空格等！</td>
                    </tr>
                    <tr> 
                      <td height="20" bgcolor="#EEEEEE" width="454" colspan="4">
                      <p align="left">　</td>
                    </tr>
                  </table>
                </td>
                <td width="17" background="../images/obj_waku3_04.gif">
                　</td>
              </tr>
            </table>
  </center>
</div>
<%End if
Call CloseRs(rs)
Call CloseDB(conn)%>
