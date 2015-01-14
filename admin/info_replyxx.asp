<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<!--#include file=../config.asp-->
<!--#include file="../Include/err.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<html>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>竞价信息操作</title>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
</head>

<body topmargin="0" leftmargin="0" background="img/bg1.gif">
<div align="center">
  <center>
  <table width="550" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
    <tr>
      <td width="214" valign="top" background="img/gggbg.jpg"><div align="center">
      </div></td>
      <td width="550" align="right" valign="top" bgcolor="#FFFFFF"><table width="550"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><form action="info_replyxx_save.asp" name="myform" method="POST" >
            <table width="100%" height="100%" border="0" align="right" cellpadding="6" cellspacing="0">
                <tr>
                  <td height="10" colspan="4"  align="center">取消到期竞价信息资格<br>(此操作将把到期竞价信息的资格及竞价金额去除，降为一般信息)</td>
                </tr> 			
                    <tr> 
                        <td height="25" align="center">
                        取消资格：<input type="radio" value="1" name="hfxxzg">是</font><input type="radio" value="0" name="hfxxzg" checked>否</td> 
                  <td height="1" colspan="2"></td>
                </tr>  
                <tr>
                  <td height="25" colspan="2"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><div align="center">
                            <input name="I2" type="submit" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #a2a2a2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #a2a2a2 1px solid; BORDER-TOP: #ffffff 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;"  value="确定" border="0">
                        </div></td>
                        <td><div align="center">
                            <input name="I2" type="reset" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #a2a2a2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #a2a2a2 1px solid; BORDER-TOP: #ffffff 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;"  value="取消" border="0">
                        </div></td>
                      </tr>
                     </tr>

                  </table></td>
                </tr>
              </table>
          </form></td>
        </tr>
      </table></td>
      <td width="4" align="right" valign="top" bgcolor="#e6e6e6"></td>
    </tr>

  </table>
  </center>
</div>
</body>
</html>
