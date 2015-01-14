
<!--#include file=cookies.asp-->
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
<title>COAD System</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../css/style.css" rel="stylesheet" type="text/css">
<script language=javascript src=../Include/mouse_on_title.js></script>
<style type="text/css">
<!--
.style1 {color: #FFFFFF;
	font-weight: bold;
}
-->
</style>
</head>
<title>COAD System</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language=javascript src=../Include/mouse_on_title.js></script>
<link href="../include/css.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {color: #FFFFFF;
	font-weight: bold;
}
-->
</style>
</head>
<body>
<%Dim softurl
softurl=Request.Servervariables("server_name")&replace(Request.Servervariables("url"),"/code.asp","")

%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#6699cc">
  <tr><td height="28" colspan="2"><div align="center" ><font color="#FFFFFF"><b>广告代码示例</b></font></div></td>
 
  </tr>
  <tr>
    <td height="250" align="center" bgcolor="#FFFFFF"> 纯JS文件（需要使用“生成JS”功能后才可使用）<br>
        <br>		<textarea name="S1" cols="80" rows="6" class="input2"><!--广告代码开始，主要用于静态页面-->
  <script src="ads/JS/广告ID号_一级地区ID号_二级地区ID号_三级地区ID号.js"></script>
  <!--广告代码结束-->
      </textarea>

        <br>
		JS代码（JS动态调用方式）<br>
        <textarea name="S1" cols="80" rows="6" class="input2"><!--广告代码开始，用于动态页面-->
  <script src="&lt;%=adjs_path("ads/js","广告ID号",c1&"_"&c2&"_"&c3)%>"></script>
  <!--广告代码结束-->
      </textarea>
        <br>
        <br>
      JS代码（ASP调用方式）<br>
      <br>
      <textarea name="textarea" cols="80" rows="6" class="input2"><!--广告代码开始。用于动态页面-->
  <script src="ads/ad.asp?adid=广告ID号&c=&lt;%=c1%>,&lt;%=c2%>,&lt;%=c3%>"></script>
<!--广告代码结束-->
      </textarea></td> 
  </tr>
  <tr>
    <td height="80" bgcolor="#FFFFFF"><table width="98%"  border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td>将上面代码的“广告ID号”更改成相应广告的ID。<br>
		将上面的地区ID号更改成相应地区的ID号<br>
*推荐! 纯JS文件代码，不需要服务器执行，速度第一。缺点不统计显示次数。<br>
*对于普通固定显示的图片/FLASH广告，将上面的代码复制到要显示广告的位置。<br>
*对于其他类型广告复制到显示页面的任意位置即可。</td>
      </tr>
    </table></td>
  </tr>
</table>
<br>
</body>
</html>