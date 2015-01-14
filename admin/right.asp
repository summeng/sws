<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/Version.asp"-->
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
<title>管理员页面</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
body {
	font-size:9pt;
	background-color: #E3EBF9;
}
td  { font-size: 9pt  }
.style2 {
	color: #FFFFFF;
	font-weight: bold;
}
.style3 {color: #FF0000}
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
}
a:hover {
	text-decoration: none;
}
a:active {
	text-decoration: none;
}
.table{BORDER: #3F8805 1px solid;background-color:#3F8805;margin-left:12px}
tr{BACKGROUND-COLOR: #EEFEE0}
.backs{BACKGROUND-COLOR: #3F8805;COLOR: #ffffff;text-align:center}
-->
</style>

<base target="_self">
</head>
<BODY>
<body text="#000000">
<table width="90%" >

<tr><td width="10%"><b>官方公告：</b></a></td><td><iframe scrolling="no" frameborder="0" marginheight="0" marginwidth="0" width="100%" height="15" src="ttscunion.htm"></iframe></td></tr>

</table>
<table width="100%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#ADAED6">
<TR>
    <TD height="20" bgcolor="#799AE1" align="center" colspan="13"><FONT color="#FFFFFF" style="font-size:14px">管 理 首 页</FONT></TD>
 </TR>
  <tr>
    <td height="28" bgcolor="#BDBEDE"><span style="font-size: 9pt"><font size="3"><font size="2">&nbsp;</font></font></span><%=title%><b>管理说明</b></td>
  </tr>
  <tr>
    <td><table width="100%"  border="0" cellpadding="6" cellspacing="1" bgcolor="#eeeeee">
      <tr>
        <td  width="60"><span class="style2"><span style="font-size: 9pt"><font size="3"></font></span></span>用&nbsp;&nbsp;&nbsp;&nbsp;户：</td>
        <td>用户向天天网络科技工作室购买本程序后，拥有《最终用户许可协议》规定下的使用权。用户具有使用本软件建立的网站的所有管理权限。 第一次使用时请到[<A href="add_administrator.asp">管理帐号设置</A>]-[增加一个新的管理员（要设置最高权限）]之后退出以新管理员登录删除系统默认建立的[admin]管理帐号，注意，不要使用简单的数字或单词作为密码，建议用别人不易猜得到的数字加字符，且不要有特定意思。以防别猜到密码登录破坏。 </td>
      </tr>
      <tr>
        <td width="60">使用设置：</td>
        <td >请不要让其他不了解该系统或是无关的人管理系统，以免破坏数据库，使数据库数据丢失。</td>
      </tr>
      <tr>
        <td width="60">注意事项：</td>
        <td>由于系统是采用COOKIE进行管理员验证的，所以每次登陆管理过后，请正常点击左上角的[<span class="style3"><a href="exit.asp">安全退出</a></span>]，以免黑客或居心不良的人从你的COOKIE中获得管理密码！</td>
      </tr>
    </table></td>
  </tr>
</table>
<center>
<br>
  <table width="100%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#ADAED6">
    <tr>
      <td height="28" bgcolor="#BDBEDE"><span style="font-size: 9pt"><font size="3"><font size="2">&nbsp;</font></font></span><b>版权声明</b></td>
    </tr>
    <tr>
      <td><table width="100%"  border="0" cellpadding="6" cellspacing="1" bgcolor="#eeeeee">

          <tr>
            <td width="60">版权声明：</td>
            <td>
              <p>1、本软件是由天天科技开发的，版权归<a href="http://www.ip126.com/" target="_blank"><FONT color=#0000ff>天天网络科技工作室</FONT></a>所有；<BR>
              2、本软件受中华人民共和国《著作权法》《计算机软件保护条例》等相关法律、法规保护，作者保留一切权利。 <BR>
              3、正式版本不准非法拷备、出售、转让，否则一切后果自负！正版用户只能一套程序在一个域名下安装，不得多域名安装，非法转让、出售、赠予和交给他人研究，否则停止相关服务并追究法律责任。<br>
			  4、用户必须无条件接受本软件的《软件使用协议》。该协议文档可到<a href="http://www.ip126.com/" target="_blank"><FONT color=#0000ff>客服中心</FONT></a>－最新下载栏目下载阅读。更多帮助请看附带的说明书。</p></td>
          </tr>
      </table></td>
    </tr>
  </table>
  <center>
<br>
    <table width="100%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#ADAED6">
    <tr>
      <td height="28" bgcolor="#BDBEDE"><span style="font-size: 9pt"><font size="3"><font size="2">&nbsp;</font></font></span><b>法律声明</b></td>
    </tr>
    <tr>
      <td><table width="100%"  border="0" cellpadding="6" cellspacing="1" bgcolor="#eeeeee">

          <tr>
            <td width="60">法律声明：</td>
            <td>
              <p>1、同大多数互联网软件一样，易受到各种安全问题的困扰，用户使用“本软件”涉及到互联网各种因素，可能会受到各个环节不稳定因素的影响，存在因不可抗力、计算机病毒、黑客攻击、系统不稳定、用户所在位置以及其他任何网络、技术、通信线路等原因造成的软件运行中断、数据丢失、信息泄露、造成损失或不能满足用户要求的风险，用户须明白并自行承担以上风险。<br>
2、用户应对利用本软件建立网站之网页所登载内容的真实性、准确性、合法性负全面责任，一切由此所引起的纠纷、争议及所涉及的法律责任均由用户承担。<br>
3、用户必须遵守国家和当地有关法律、法规、行政规章，不得利用本软件制作、复制、发布、传播任何法律法规禁止的有害信息。用户对其经营行为和发布的信息违反有关规定而引起的任何的政治责任、法律责任要承担全部责任，与软件作者无关。</p></td>
          </tr>
      </table></td>
    </tr>
  </table>
  <br>
   <table width="100%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#ADAED6">
    <tr>
      <td height="28" bgcolor="#BDBEDE"><span style="font-size: 9pt"><font size="3"><font size="2">&nbsp;</font></font></span><b>售后服务</b></td>
    </tr>
    <tr>
      <td><table width="100%"  border="0" cellpadding="6" cellspacing="1" bgcolor="#eeeeee">
	  	   <tr>
            <td width="60"><span class="style2"><span style="font-size: 9pt"><font size="3"></font></span></span>当前版本：</td>
            <td><%Response.Write "<A HREF='http://www.ip126.com' TARGET='_blank'>TianTian &reg; INFO of SAD&#8482; " & SystemVersion & " "&DatabaseType &" Build " & SystemBuildDate & "</a>"%></td>
          </tr>
          <tr>
            <td width="60"><span class="style2"><span style="font-size: 9pt"><font size="3"></font></span></span>升级说明：</td>
            <td>正版用户根据不同的价位获得不同的服务。以<a href="http://www.ip126.com/" target="_blank">官方客服中心网站</a>为准</td>
          </tr>
         <tr>
            <td width="60">联系方式：</td>
            <td><p>
			官方客服中心网站：<a href="http://www.ip126.com/" target="_blank">http://www.ip126.com/</a><BR>
              电子邮件：bdunion@126.com<BR>
              腾讯QQ：530051328（加好友注明天天源码）<BR>
            </p></td>
          </tr>
      </table></td>
    </tr>
  </table>

    <br>
  <table width="100%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#ADAED6">
    <tr>
      <td height="28" bgcolor="#BDBEDE"><span style="font-size: 9pt"><font size="3"><font size="2">&nbsp;</font></font></span><b>基本组件检测</b> （如果下面五项都是<font color=green><b>√</b></font>即支持我们的系统）</td>
    </tr>
</table>
<%
Function IsObjInstalled(strClassString)'定义函数
 On Error Resume Next'有错则跳过
 IsObjInstalled = False'函数名初始值设为FALSE
 Err = 0'错误计数器置零
 Dim xTestObj'定义对象
 Set xTestObj = Server.CreateObject(strClassString)'初始化对象
 If 0 = Err Then IsObjInstalled = True'如果错误计数器为0 则函数名为真 意思是已经安装了
 Set xTestObj = Nothing'销毁对象
 Err = 0'错误计数器置零
End Function'结束函数

Function GetVer(ClassStr)
 On Error Resume Next
 GetVer = ""
 Err = 0
 Dim xTestObj
 Set xTestObj = Server.CreateObject(ClassStr)
 If 0 = Err Then GetVer = xTestObj.Version
 Set xTestObj = Nothing
 Err = 0
End Function

Sub WRITE_LINE(str)
 Response.Write LTrim(str)'输出str 去除左边空格的
End Sub

Dim InstallObj(4)'检查各个组件 0-3 意思自己百度
InstallObj(0) = "adodb.connection"
InstallObj(1) = "Scripting.FileSystemObject"
InstallObj(2) = "Persits.Jpeg"
InstallObj(3) = "JMail.Message"

Dim name(4)'各个组件简称
name(0) = "Access"
name(1) = "FSO"
name(2) = "ASPJpeg"
name(3) = "JMail"

Dim Effect(4)'各个组件的作用
Effect(0) = "使用Access存取数据"
Effect(1) = "FSO 文件系统管理、文本文件读写、文件生成或删除"
Effect(2) = "系统通过ASPJpeg组件生成图片标题、水印、阻止伪装图片木马上传等"
Effect(3) = "系统通过JMail组件发送电子邮件"


Dim Suggestion(4)'建议
Suggestion(0) = "联系空间商解决Access数据库问题"
Suggestion(1) = "找支持FSO的空间"
Suggestion(2) = "服务器要安装ASPJpeg组件，或找支持ASPJpeg的空间，否则无法上传图片。"
Suggestion(3) = "服务器要安装JMail组件，或找支持JMail的空间，否则无法发送邮件"
%> 
    <tr>
      <td><table border=0 width="100%" cellspacing=1 cellpadding=3>
  <tr class="backs"><td width="200">组件名称</td> <td width="100">简称</td> <td width="400" >作 用</td><td width="80">支持/版本</td><td width="100">建议</td></tr>
  <tr><td width="200">Active Server Page</td> <td width="100">ASP</td> <td width="400" >与数据库和其它程序进行交互</td><td width="80"><%response.write "<font color=green><b>√</b></font>"%></td><td width="100">已支持</td></tr> 

<%Dim i,FoundFso
i=0
for i=0 to 3%>
  <tr><td width="200"><%=InstallObj(i)%></td>
  <td width="100" ><%=name(i)%></td>
  <td width="400" ><%=Effect(i)%></td>
   <td width="80" ><%
FoundFso = IsObjInstalled(InstallObj(i))
 
If FoundFso Then
 Response.Write "<font color=green><b>√</b></font>"&GetVer(InstallObj(i))&""
Else
 Response.Write "<font color=red><b>×</b></font>"
End If
%></td>
<td width="100"><%If FoundFso Then
Response.Write "已支持"
Else
 Response.Write Suggestion(i)
End If
%></a>
  </tr>
 <%
 next%>
	  <tr>
		<td>IIS版本</td><td colspan="4"><%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
	  </tr>
   <tr><td width="200" colspan="5"><a href="admin_server.asp" TARGET=right>更多项目检测>>></a></td>
  </tr>
 </table></td>
    </tr>
  </table>

</center>
</body>
</html>