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
dim username,id
id=trim(request("id"))
if not isnumeric(id) or id="" then
response.write "<li>参数错误！"
response.end
end if
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select username,sdname,sdjj,sdfg,sdkg,mpname,mpjj,mpfg,mpkg,dpdata,http,zhuying,city_oneid,city_twoid,city_threeid,type_oneid,type_twoid,type_threeid,sdgd,mpgd,notification from [DNJY_user] where id="&cstr(id) 
rs.open sql,conn,1,1
%>
<meta http-equiv="Content-Language" content="zh-cn">
<title>修改用户店铺名片资料</title>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<body topmargin="0" leftmargin="0">
<SCRIPT language=javascript>
<!--
//字节数限制，与下面配合=========
String.prototype.Length=function(){
temp=this.replace(/([^\x00-\xff])/g,"$1$1");
return temp.length
}
function CheckForm()
{

     if (document.thisForm.sdname.value.Length()>30) //字节数限制，与上面配合
     {
	 alert("网店名称长度要求30字节内，请重新输入！");
	  document.thisForm.sdname.focus()
	  return false
  }	 
if ((document.thisForm.sdname.value.length>1) && (document.thisForm.sdjj.value==""))
	{
	    alert("请填网店简介!");
		document.thisForm.sdjj.focus();
	    return false;
	}
	if ((document.thisForm.sdname.value.length>1) && (document.thisForm.sdfg.value=="0" ))
	{
	    alert("请选择网店风格");
		document.thisForm.sdfg.focus();
	    return false;
	}
     if (document.thisForm.mpname.value.Length()>30) //字节数限制，与上面配合
     {
	 alert("名片名称长度要求30字节内，请重新输入！");
	  document.thisForm.mpname.focus()
	  return false
  }	 
if ((document.thisForm.mpname.value.length>1) && (document.thisForm.mpjj.value==""))
	{
	    alert("请填名片简介!");
		document.thisForm.mpjj.focus();
	    return false;
	}
if ((document.thisForm.mpname.value.length>1) && (document.thisForm.mpfg.value=="0" ))
	{
	    alert("请选择名片风格");
		document.thisForm.mpfg.focus();
	    return false;
    }
if (document.thisForm.type_one.value=="")
	{
	  alert("请选择行业分类！")
	  document.thisForm.type_one.focus()
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
if (document.thisForm.notification.value.length>200)
	{
	  alert("审核通知内容不能超过200个字符")
	  document.thisForm.notification.focus()
	  return false
	 }	
}

//-->
     </SCRIPT>
<div align="center">
  <center><form method="POST" name="thisForm" action="user_editdp_save.asp?id=<%=id%>" onSubmit="return CheckForm();">
            <table width="489" border="0" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
              <tr> 
                <td width="16" background="../images/obj_waku3_03.gif">
                　</td>
                <td width="456" bgcolor="#EEEEEE">
                  <table width="454" border="0" cellspacing="0" cellpadding="0" height="177">
                    <tr>
                      <td height="22" bgcolor="#EEEEEE" colspan="4">
                      <p align="center"><b>修改会员店铺和名片资料</b></td>
                    </tr>
                    <tr>
                      <td height="30" bgcolor="#EEEEEE" colspan="4"><HR></td>
                    </tr>
                    <tr>
                      <td height="30" bgcolor="#E3EBF9" colspan="2">
                      <p align="right">网店名称：</td>
                      <td height="30" bgcolor="#E3EBF9" colspan="2">
                      &nbsp;
                      <input name="sdname" type="text" id="sdname" value="<%=rs("sdname")%>" size="30" maxlength="17"></td>
                    </tr>

 
                    <tr>
                      <td height="30" bgcolor="#E3EBF9" colspan="2" align="right">网店风格：</td>
                      <td height="30" bgcolor="#E3EBF9" colspan="2">&nbsp;
                          <select name="sdfg">
						   <option  value="0" style="background-color:#F8F4F5 ;color: #FF0000">选择网店风格</option>
                            <option  value="1" <%if rs("sdfg")=1 then%>selected<%end if%>>蓝天白云</option>
                            <option  value="2" <%if rs("sdfg")=2 then%>selected<%end if%>>黄金季节</option>
							<option  value="3" <%if rs("sdfg")=3 then%>selected<%end if%>>绿草如茵</option>
							<option  value="4" <%if rs("sdfg")=4 then%>selected<%end if%>>万紫千红</option>	
                          </select>
 &nbsp;&nbsp; <a href="../img/dqmb.jpg" target="_blank">效果预览：</a></td>
                    </tr>
                    <tr> 
                      <td height="30" bgcolor="#E3EBF9" colspan="2">
                      <p align="right">网店简介：</td>
                      <td height="30" bgcolor="#E3EBF9" colspan="2">
                      &nbsp;
                      <label>
                      <textarea name="sdjj" cols="50" rows="8" id="sdjj"><%=rs("sdjj")%></textarea>
                      </label></td>
                    </tr>
<tr>
                      <td height="30" bgcolor="#E3EBF9" colspan="2" align="right">网店审核：</td>
                      <td height="30" bgcolor="#E3EBF9" colspan="2">&nbsp;
                          <select name="sdkg">
                            <option value="0" <%if rs("sdkg")=0 then%>selected<%end if%>>关闭状态</option>
                            <option value="1" <%if rs("sdkg")=1 then%>selected<%end if%>>开放状态</option>
			    <option value="2" <%if rs("sdkg")=2 then%>selected<%end if%>>申请开放</option>
                                                    </select>
当前状态<%
if rs("sdkg")=1 then
response.write "<font color=""#ff0000""><strong>开放状态</strong></font>"
elseif rs("sdkg")=0 then
response.write "<font color=""#0066FF""><strong>关闭状态</strong></font>"
elseif rs("sdkg")=2 then
response.write "<font color=""#FF9900""><strong>申请开放</strong></font>"
end if%></td>
                    </tr>
<tr>
                      <td height="30" bgcolor="#E3EBF9" colspan="2" align="right">网店固顶：</td>
                      <td height="30" bgcolor="#E3EBF9" colspan="2">&nbsp;
                          <select name="sdgd">
                            <option value="0" <%if rs("sdgd")=0 then%>selected<%end if%>>未固顶</option>
                            <option value="1" <%if rs("sdgd")=1 then%>selected<%end if%>>已固顶</option>
							</select>
</td>
                    </tr>
                    <tr>
                      <td height="30" bgcolor="#EEEEEE" colspan="4"><HR></td>
                    </tr>					
                    <tr>
                      <td height="30" bgcolor="#CCCCFF" colspan="2">
                      <p align="right">名片名称：</td>
                      <td height="30" bgcolor="#CCCCFF" colspan="2">
                      &nbsp;
                      <input name="mpname" type="text" id="mpname" value="<%=rs("mpname")%>" size="30" maxlength="17"></td>
                    </tr>
 
                    <tr>
                      <td height="30" bgcolor="#CCCCFF" colspan="2" align="right">名片风格：</td>
                      <td height="30" bgcolor="#CCCCFF" colspan="2">&nbsp;
                          <select name="mpfg">
						   <option  value="0" style="background-color:#F8F4F5 ;color: #FF0000">选择名片风格</option>
                            <option  value="1" <%if rs("mpfg")=1 then%>selected<%end if%>>蓝天白云</option>
                            <option  value="2" <%if rs("mpfg")=2 then%>selected<%end if%>>深蓝宝石</option>
							<option  value="3" <%if rs("mpfg")=3 then%>selected<%end if%>>青青河草</option>
							<option  value="4" <%if rs("mpfg")=4 then%>selected<%end if%>>黄金季节</option>
							<option  value="5" <%if rs("mpfg")=5 then%>selected<%end if%>>温馨粉红</option>
                          </select>
 &nbsp;&nbsp; <a href="../img/dqmb.jpg" target="_blank">效果预览：</a></td>
                    </tr> 
                    <tr> 
                      <td height="30" bgcolor="#CCCCFF" colspan="2">
                      <p align="right">名片简介：</td>
                      <td height="30" bgcolor="#CCCCFF" colspan="2">
                      &nbsp;
                      <label>
                      <textarea name="mpjj" cols="50" rows="8" id="mpjj"><%=rs("mpjj")%></textarea>
                      </label></td>
                    </tr>
<tr>
                      <td height="30" bgcolor="#CCCCFF" colspan="2" align="right">名片审核：</td>
                      <td height="30" bgcolor="#CCCCFF" colspan="2">&nbsp;
                          <select name="mpkg">
                            <option value="0" <%if rs("mpkg")=0 then%>selected<%end if%>>关闭状态</option>
                            <option value="1" <%if rs("mpkg")=1 then%>selected<%end if%>>开放状态</option>
			                <option value="2" <%if rs("mpkg")=2 then%>selected<%end if%>>申请开放</option>
                                                    </select>
当前状态<%
if rs("mpkg")=1 then
response.write "<font color=""#ff0000""><strong>开放状态</strong></font>"
elseif rs("mpkg")=0 then
response.write "<font color=""#0066FF""><strong>关闭状态</strong></font>"
elseif rs("mpkg")=2 then
response.write "<font color=""#FF9900""><strong>申请开放</strong></font>"
end if%></td>
                    </tr>
<tr>
                      <td height="30" bgcolor="#CCCCFF" colspan="2" align="right">名片固顶：</td>
                      <td height="30" bgcolor="#CCCCFF" colspan="2">&nbsp;
                          <select name="mpgd">
                            <option value="0" <%if rs("mpgd")=0 then%>selected<%end if%>>未固顶</option>
                            <option value="1" <%if rs("mpgd")=1 then%>selected<%end if%>>已固顶</option>
							</select>
</td>
                    </tr>
                    <tr>
                      <td height="30" bgcolor="#EEEEEE" colspan="4"><HR></td>
                    </tr>
<tr>
                  <td height="30" bgcolor="#EEEEEE" colspan="2"><p align="right">所属行业类别：</td>
                  <td height="30" bgcolor="#EEEEEE" colspan="2">
<%Dim rsi
set rsi=conn.execute("select * from DNJY_hy_type where id>0 and twoid>0 and threeid=0 order by indexid")%>
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
                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" colspan="2">
                      <p align="right">主 &nbsp;营：</td>
                      <td height="30" bgcolor="#EEEEEE" colspan="2">
                      &nbsp;
                      <label>
                      <input name="zhuying" type="text" class="inputa" value="<%=rs("zhuying")%>" size="50" maxlength="50">
                      </label></td>
                    </tr>

                    <tr>
                      <td height="30" bgcolor="#EEEEEE" colspan="2">
                      <p align="right">公司网站：</td>
                      <td height="30" bgcolor="#EEEEEE" colspan="2">
                      &nbsp;
                      <input name="http" type="text" class="inputa" id="http" value="<%=rs("http")%>" size="30" maxlength="50"> 前面不能带http://</td>
                    </tr>            					
                     <tr>
                      <td height="30" bgcolor="#EEEEEE" colspan="4"><HR></td>
                    </tr>                   
                      <tr> 
                      <td height="30" bgcolor="#ffffff" colspan="2">
                      <p align="right">审核事项通知：</td>
                      <td height="30" bgcolor="#ffffff" colspan="2">
                      &nbsp;
                      <label>
                      <textarea name="notification" cols="50" rows="5" id="notification"><%=rs("notification")%></textarea>
                      </label>(限200字符)</td>
                    </tr>                  
                    <tr>
                      <td height="30" bgcolor="#ffffff" colspan="2">
                      <p align="right">是否更新审核时间：</td>
                      <td height="30" bgcolor="#ffffff" colspan="2">
                      &nbsp;
                      <input type="radio" name="gxsj" value="1" id="switch" >
                 更新&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="gxsj" value="0" id="switch" checked>不更新<br></td>
                    </tr>                        
                    <tr>
                      <td height="30" bgcolor="#ffffff" colspan="2">
                      <p align="right">审核时间：</td>
                      <td height="30" bgcolor="#ffffff" colspan="2">
                      &nbsp;
                      <%=rs("dpdata")%></td>
                    </tr>
					<input type="hidden" name="username" value="<%=rs("username")%>">
                    <tr> 
                      <td height="25" bgcolor="#ffffff" width="110">
                      <p align="center">　</td>
                      <td height="25" bgcolor="#ffffff" colspan="2">
                      <p align="right">
                      　</td>
                      <td height="25" bgcolor="#ffffff" width="280">
                      <p align="center">
                      <input border="0" src="../images/Login_but.gif" name="I4" type="image">
                &nbsp;&nbsp;&nbsp;&nbsp;<a href="user_certificate.asp?id=<%=cstr(id)%>" target=blank><font color=#0000ff>查看其上传的证照</font></a></td>
                    </tr>
                    <tr> 
                      <td height="10" bgcolor="#ffffff" colspan="4">&nbsp;</td>
                    </tr>
                  </table>
                </td>
                <td width="17" background="../images/obj_waku3_04.gif"></td>
              </tr>
            </table></form>
  </center>
</div>
<%Call CloseRs(rs)
Call CloseDB(conn)%>
