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
<%Dim problem
Call OpenConn
problem=trim(conn.Execute("Select problem From DNJY_user Where username='"&request.cookies("dick")("username")&"'")(0))
Call CloseDB(conn)%>
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
<link rel="stylesheet" type="text/css" href="../1.CSS">
<style type="text/css">
<!--
.style1 {color: #FF0000}
.style5 {color: #FF0000; font-weight: bold; }
-->
</style>
<script language="javascript">
<!--
//字节数限制，与下面配合=========
String.prototype.Length=function(){
temp=this.replace(/([^\x00-\xff])/g,"$1$1");
return temp.length
}
// --></script>
<script language="javascript" type="text/javascript">
<!--
function CheckForm()
{

        if(document.thisForm.pass1.value.length<1)
	{
	    alert("原密码不能为空!");
		document.thisForm.pass1.focus()
	    return false;
	}
	  if ((document.thisForm.pass2.value.Length()>12) || (document.thisForm.pass2.value.Length()<5)) //字节数限制，与上面配合
     {
	 alert("新密码长度要求5－12字节，请重新输入！");
	  document.thisForm.pass2.focus()
	  return false
  }
	    if(document.thisForm.pass3.value.length<1)
	{
	    alert("确认新密码不能为空!");
		document.thisForm.pass3.focus()
	    return false;
	}
	   if(document.thisForm.pass3.value!=document.thisForm.pass2.value)
        {
            alert("两次输入密码不一致!");
			document.thisForm.pass3.focus()
            return false;
        }
}


  function CheckForm_answer()
    {
<%if IsNull(problem) then%>
        if(document.Form.problem.value.length<1)
	{
	    alert("密码问题不能为空!");
		document.Form.problem.focus()
	    return false;
	}
        if(document.Form.answer1.value.length<1)
	{
	    alert("答案不能为空!");
		document.Form.answer1.focus()
	    return false;
	}
	    if(document.Form.answer2.value.length<1)
	{
	    alert("确认答案不能为空!");
		document.Form.answer2.focus()
	    return false;
	}
	   if(document.Form.answer1.value!=document.Form.answer2.value)
        {
            alert("两次输入答案不一致!");
			document.Form.answer1.focus()
            return false;
        }
<%else%>
        if(document.Form.problem.value.length<1)
	{
	    alert("新密码问题不能为空!");
		document.Form.problem.focus()
	    return false;
	}
        if(document.Form.answer1.value.length<1)
	{
	    alert("原答案不能为空!");
		document.Form.answer1.focus()
	    return false;
	}
	    if(document.Form.answer2.value.length<1)
	{
	    alert("确认新答案不能为空!");
		document.Form.answer2.focus()
	    return false;
	}
	   if(document.Form.answer3.value!=document.Form.answer2.value)
        {
            alert("两次输入答案不一致!");
			document.Form.answer3.focus()
            return false;
        }
<%end if%>

}
//-->
</SCRIPT>
</head>

<body background="img/bg1.gif" leftmargin="0" topmargin="0">

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" height="426">
    <tr>
      <td width="214" height="354" valign="top" bgcolor="#FFFFFF"><div align="center">
      
        <!--#include file=userleft.asp--></div></td>
		<td width="5">&nbsp;</td>
      <td width="760" height="354" valign="top" bgcolor="#FFFFFF">        <div align="right"></div>
        <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse" class="ty1">
          
            <tr>
              <td height="25" bgcolor="#FAFAFA"><span class="style5">用户密码和密码保护问答修改：</span></td>
              <td bgcolor="#FAFAFA">&nbsp;</td>
            </tr>
            <tr bgcolor="#CCCCCC">
              <td width="456" height="1"></td>
              <td width="17"></td>
            </tr>
            <tr>
              <td width="500" valign="top" align="center">
			  <form name="thisForm"  method="POST" action="user_pass_save.asp?key=pass"  language="javascript" onsubmit="return CheckForm()">
			  <table border="0" cellspacing="0" cellpadding="0" height="212">
                  <tr>
                    <td height="22" colspan="3"><p align="center"><b>修改用户登录密码</b></td>
                  </tr>
                  <tr>
                    <td height="30"><p align="right">输入原密码：</td>
                    <td  height="30">&nbsp;
                        <input class="inputa"　type="text" name="pass1" size="20" maxlength="12"></td>
                  </tr>
                  <tr>
                    <td height="30"><p align="right">输入新密码：</td>
                    <td height="30">&nbsp;
                        <input class="inputa" type="password" name="pass2" size="20"></td>
                  </tr>
                  <tr>
                    <td height="30"><p align="right"> 确认新密码：</td>
                    <td height="30">&nbsp;
                        <input class="inputa" type="password" name="pass3" size="20"></td>
                  </tr>

                  <tr>
                    <td height="25" colspan="3" align="center"><table  width="200" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td><div align="center">
                              <input name="B1" type="submit" class="inputb" value="提交修改" border="0"   onclick="javascript:return CheckForm();">
                          </div></td>
                          <td><div align="center">
                              <input name="B2" type="reset" class="inputb" id="B2" value="取消修改" border="0" >
                          </div></td>
                        </tr>
                    </table></td>
                    <td height="25" width="66"> 　　　　　 </td>
                  </tr>
                  <tr>
                    <td height="20" colspan="3"><div align="center">请填写正确的旧密码，否则你的密码将<span class="style1">无法修改</span>！</div></td>
                  </tr>

              </table>
			 </form>
<TABLE>
<TR>
	<TD height="22"></TD>
</TR>
</TABLE>
			 <form name="Form"  method="POST" action="user_pass_save.asp?key=answer"  language="javascript" onsubmit="return CheckForm_answer()">
                 <table border="0" cellspacing="0" cellpadding="0" height="212">

				  <%If IsNull(problem) then%>
                  <tr>
                    <td height="22" colspan="3"><p align="center"><b>您尚未设置密码保护，请及时设置</b></td>
                  </tr>
                  <tr>
                    <td height="30"><p align="right">输入密保问题：</td>
                    <td  height="30">&nbsp;
                        <input class="inputa"　type="text" name="problem" size="20" maxlength="12"></td>
                  </tr>
                  <tr>
                    <td height="30"><p align="right">输入密保答案：</td>
                    <td height="30">&nbsp;
                        <input class="inputa" type="text" name="answer1" size="20"></td>
                  </tr>
                  <tr>
                    <td height="30"><p align="right"> 确认密保答案：</td>
                    <td height="30">&nbsp;
                        <input class="inputa" type="text" name="answer2" size="20"></td>
                  </tr>
				  <%else%>
                  <tr>
                    <td height="22" colspan="3"><p align="center"><b>修改密码保护资料</b></td>
                  </tr>
                  <tr>
                    <td height="30"><p align="right">密保问题：</td>
                    <td  height="30">&nbsp;<%=problem%></td>
                  </tr>
                  <tr>
                    <td height="30"><p align="right">输入新问题：</td>
                    <td  height="30">&nbsp;
                        <input class="inputa"　type="text" name="problem" size="20" maxlength="12"></td>
                  </tr>
                  <tr>
                    <td height="30"><p align="right">输入原答案：</td>
                    <td  height="30">&nbsp;
                        <input class="inputa"　type="text" name="answer1" size="20" maxlength="12"></td>
                  </tr>
                  <tr>
                    <td height="30"><p align="right">输入新答案：</td>
                    <td height="30">&nbsp;
                        <input class="inputa" type="text" name="answer2" size="20"></td>
                  </tr>
                  <tr>
                    <td height="30"><p align="right"> 确认新答案：</td>
                    <td height="30">&nbsp;
                        <input class="inputa" type="text" name="answer3" size="20"></td>
                  </tr>
				  <input type="hidden" name="Modify" value="ok">
                <%End if%>
                  <tr>
                    <td height="25" colspan="3" align="center"><table  width="200" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td><div align="center">
                              <input name="B1" type="submit" class="inputb" value="确认" border="0"   onclick="javascript:return CheckForm_answer();">
                          </div></td>
                          <td><div align="center">
                              <input name="B2" type="reset" class="inputb" id="B2" value="取消" border="0" >
                          </div></td>
                        </tr>
                    </table></td>
                    <td height="25" width="66"> 　　　　　 </td>
                  </tr>
              </table>
           </form>
			  </td>
              <td width="17"> 　</td>
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