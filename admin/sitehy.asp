<!--#include file=../config.asp-->
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>

<link href="../css/style.css" rel="stylesheet" type="text/css">
<BODY>
<style type="text/css">
<!--
.style1 {color: #FFFFFF;
	font-weight: bold;
}
.style2 {color: #FF0000}
-->
</style>

<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%">
    <TR>
    <TD height="20" bgcolor="#799AE1" align="center" colspan="13"><FONT color="#FFFFFF" style="font-size:14px">远 程 调 用 网 店 名 片 代 码 及 示 例</FONT></TD>
 </TR>
    <tr bgcolor="#BDBEDE">
      <td height="28" style="border-bottom-style: solid; border-bottom-width: 1"><span class="style1"><span style="font-size: 9pt"><span style="font-size: 9pt"><font size="3"><font size="2">&nbsp;</font></font>远程调用网店说明</span></span></span></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td style="border-top-style: solid; border-top-width: 1"><table width="98%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td><p>可在站内或站外任何页面直接使用下面代码调用，所调用的网店是动态的，无需手工生成：</p>
           
            <p class="style2">
              <textarea name="textarea" cols="100" rows="6"><!--代码开始--><script language="javascript" src="http://<%=web%>/<%=strInstallDir%>user/newsd.asp?c1=0&c2=0&c3=0&t1=0&t2=0&t3=0&gdxx=0&btw=24&s=10&l=1&zt=13&hg=10"></script><!--代码结束-->
              </textarea>
            </p>调用参数说明（顺序不能乱，参数不能超出范围）：<br>
			依次：依次：c1一级地区分类（0不分、其它按该类ID号），c2二级地区分类（0不分、其它按该类ID号），c3三级地区分类（0不分、其它按该类ID号）,t1行业一级分类（0不分、其它按该类ID号）,t2行业二级分类（0不分、其它按该类ID号）,t3行业三级分类（0不分、其它按该类ID号），gdxx更多（0不显1显），btw名称长度（以字节计），s显示条数，l列数，zt字号，hg行高
            </td>
        </tr>
      </table></td>
    </tr>
	<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="230">
   <tr>
      <td height="28" style="border-bottom-style: solid; border-bottom-width: 1">
<b>调用预览：</b>（效果是按默认的参数）
</td>
        </tr>

    <tr bgcolor="#34991B">
      <td height="28" style="border-bottom-style: solid; border-bottom-width: 1">
<!--代码开始-->
<script language="javascript" src="../user/newsd.asp?c1=0&c2=0&c3=0&t1=0&t2=0&t3=0&gdxx=1&btw=24&s=10&l=1&zt=13&hg=10"></script>
<!--代码结束-->
</td>
        </tr>
     </table></table><br>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%">
		    <tr bgcolor="#BDBEDE">
      <td height="28" style="border-bottom-style: solid; border-bottom-width: 1" ><span class="style1"><span style="font-size: 9pt"><span style="font-size: 9pt"><font size="3"><font size="2">&nbsp;</font></font>远程调用名片说明</span></span></span></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td style="border-top-style: solid; border-top-width: 1"><table width="98%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td><p>可在站内或站外任何页面直接使用下面代码调用，所调用的名片是动态的，无需手工生成：</p>
           
            <p class="style2">
              <textarea name="textarea" cols="100" rows="6"><!--代码开始--><script language="javascript" src="http://<%=web%>/<%=strInstallDir%>user/newhy.asp?c1=0&c2=0&c3=0&t1=0&t2=0&t3=0&gdxx=0&btw=24&s=10&l=1&zt=13&hg=10"></script><!--代码结束-->
              </textarea>
            </p>调用参数说明（顺序不能乱，参数不能超出范围）：<br>
			依次：依次：c1一级地区分类（0不分、其它按该类ID号），c2二级地区分类（0不分、其它按该类ID号），c3三级地区分类（0不分、其它按该类ID号）,t1行业一级分类（0不分、其它按该类ID号）,t2行业二级分类（0不分、其它按该类ID号）,t3行业三级分类（0不分、其它按该类ID号），gdxx更多（0不显1显），btw名称长度（以字节计），s显示条数，l列数，zt字号，hg行高
            </td>
        </tr>
      </table></td>
    </tr>
	<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="230">
   <tr>
      <td height="28" style="border-bottom-style: solid; border-bottom-width: 1">
<b>调用预览：</b>（效果是按默认的参数）
</td>
        </tr>

    <tr bgcolor="#34991B">
      <td height="28" style="border-bottom-style: solid; border-bottom-width: 1">
<!--代码开始-->
<script language="javascript" src="../user/newhy.asp?c1=0&c2=0&c3=0&t1=0&t2=0&t3=0&gdxx=1&btw=24&s=10&l=1&zt=13&hg=10"></script>
<!--代码结束-->
</td>
        </tr>
      </table>
</table>

