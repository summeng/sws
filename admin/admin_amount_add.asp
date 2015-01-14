<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("03")
dim username,aliname,ywje,ywlx,czbz,ddhm,zffs,dick,admincl
username=trim(request("username"))
aliname=trim(request("aliname"))
dick=trim(request("dick"))
if username="" And dick="yz" then
response.write "<li>参数错误！"
response.end
end If
If trim(request("dick"))="yz" Then
aliname=trim(request("aliname"))
ywje=Formatnumber(request("ywje"),2,-1,-1,0)
ddhm=trim(request("ddhm"))
zffs=trim(request("zffs"))
ywlx="入帐"
czbz=""&ddhm&" 订单金额入帐"
admincl=1
End If
If ddhm="" then czbz="客服操作"
%>

<meta http-equiv="Content-Language" content="zh-cn">
<title>财务操作</title>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<body topmargin="0" leftmargin="0">
<div align="center">
  <center>
            <table width="489" border="0" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
              <form method="POST" name="thisForm" action="admin_amount_add_save.asp">
              <tr> 
                <td width="16" background="../images/obj_waku3_03.gif">
                　</td>
                <td width="456" bgcolor="#EEEEEE">
                  <table width="454" border="0" cellspacing="0" cellpadding="0" height="297">
                    <tr>
                      <td height="22" bgcolor="#EEEEEE" width="454" colspan="4">
                      <p align="center">财 务 操 作<br><font color="#FF0000">（提交后将给会员<%=username%>的帐户增加金额<%=ywje%>！）</font></td>
                      </tr>
                    <tr>
                      <td height="30" bgcolor="#EEEEEE" width="454" colspan="2">
                      <p align="right">会员ID号：</td>
                      <td height="30" bgcolor="#EEEEEE" width="454" colspan="2">
                      &nbsp;<input type="text" maxlength="12" name="username" size="24" value="<%=username%>" <%If dick="yz" Or dick="uyz" then%>DISABLED<%End if%>></td>
                    </tr>
                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="454" colspan="2">
                      <p align="right">用户姓名：</td>
                      <td height="30" bgcolor="#EEEEEE" width="454" colspan="2">
                      &nbsp;<input type="text" maxlength="50" name="aliname" size="24" value="<%=aliname%>" <%If dick="yz" Or dick="uyz" then%>DISABLED<%End if%>></td>
                    </tr>
                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="454" colspan="2">
                      <p align="right">金&nbsp;&nbsp;&nbsp;&nbsp;额：</td>
                      <td height="30" bgcolor="#EEEEEE" width="454" colspan="2">
                      &nbsp;<input style="BACKGROUND-COLOR: #ffffff; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="ywje" size="10" maxlength="30" value="<%=ywje%>" <%If dick="yz" then%>DISABLED<%End if%> onkeyup="value=value.replace(/[^\d|^\.]/g,'')"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))"> 元</td>
                    </tr><%If dick="yz" then%>
                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="454" colspan="2">
                      <p align="right">支付方式：</td>
                      <td height="30" bgcolor="#EEEEEE" width="454" colspan="2">
                      &nbsp;<select size="1" value="0" name="zffs">
                <option value="0" <%If zffs="0" then%>selected<%end if%>>选择支付方式</option>
                <option value="1" <%If zffs="1" then%>selected<%end if%>>在线支付</option>
                <option value="2" <%If zffs="2" then%>selected<%end if%>>农行汇款</option>
                <option value="3" <%If zffs="3" then%>selected<%end if%>>工行汇款</option>
                <option value="4" <%If zffs="4" then%>selected<%end if%>>建行汇款</option>
                <option value="5" <%If zffs="5" then%>selected<%end if%>>农信汇款</option>
                <option value="6" <%If zffs="6" then%>selected<%end if%>>邮局汇款</option>
                </select></td>
                    </tr>
					<%End if%>
					<%If dick<>"yz" then%>
                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="454" colspan="2">
                      <p align="right">财务类型：</td>
                      <td height="30" bgcolor="#EEEEEE" width="454" colspan="2">
                      &nbsp;<select id="ywlx" name="ywlx">
                      <option <%if ywlx="入帐" then%>selected<%end if%> value="入帐">入帐</option>
                      <option <%if ywlx="支出" then%>selected<%end if%> value="支出">支出</option>
					  <option <%if ywlx="撤帐" then%>selected<%end if%> value="撤帐">撤帐</option>
                      </select></td>
                    </tr>
					<%End if%>
                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="454" colspan="2">
                      <p align="right">备&nbsp;&nbsp;&nbsp;&nbsp;注：</td>
                      <td height="30" bgcolor="#EEEEEE" width="454" colspan="2">
                      &nbsp;<textarea rows="8" name="czbz" cols="30"><%=czbz%></textarea></td>
                    </tr>
				 <%If dick="yz" Or dick="uyz" then%>
				 <input type="hidden" name="username" value="<%=username%>">
				 <input type="hidden" name="aliname" value="<%=aliname%>">
				 <%End if%>
				 <%If dick="yz" Or dick="noyz" then%>
				 <input type="hidden" name="ddhm" value="<%=ddhm%>">
				 <input type="hidden" name="ywlx" value="<%=ywlx%>">
				 <input type="hidden" name="ywje" value="<%=ywje%>">
				 <%End if%>

                      <td height="25" bgcolor="#EEEEEE" width="454">
                      <p align="center">　</td>
                      <td height="25" bgcolor="#EEEEEE" width="454" colspan="2">
                      <p align="right">
                      　</td>
                      <td height="25" bgcolor="#EEEEEE" width="454">
                      <p align="center">
					  
                      <input border="0"  src="../images/Login_but.gif" name="I4" type="image"></td>
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
