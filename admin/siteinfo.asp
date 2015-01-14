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
    <TD height="20" bgcolor="#799AE1" align="center" colspan="13"><FONT color="#FFFFFF" style="font-size:14px">远 程 调 用 代 码 及 示 例</FONT></TD>
 </TR>
    <tr bgcolor="#BDBEDE">
      <td height="28" style="border-bottom-style: solid; border-bottom-width: 1"><span class="style1"><span style="font-size: 9pt"><span style="font-size: 9pt"><font size="3"><font size="2">&nbsp;</font></font>远程调用说明</span></span></span></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td style="border-top-style: solid; border-top-width: 1"><table width="98%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td><p>可在站内或站外任何页面直接使用下面代码调用，所调用的信息是动态的，无需手工生成：</p>
           
            <p class="style2">
              <textarea name="textarea" cols="100" rows="6"><!--代码开始--><script language="javascript" src="http://<%=web%>/<%=strInstallDir%>user/newinfo.asp?xxsx=0&tt1=0&tt2=0&tt3=0&c1=0&c2=0&c3=0&xxlx=1&xxtj=0&gd=0&tp=1&xxsj=0&jj=0&gdxx=0&ys=0&hits=0&btw=20&s=10&l=1&zt=13&hg=10"></script><!--代码结束-->
              </textarea>
            </p>调用参数说明（顺序不能乱，参数不能超出范围）：<br>
			依次：xxsx信息属性（1最新、2竞价、3推荐、4热门），tt1一级信息分类（0不分、其它按该类ID号），tt2二级信息分类（0不分、其它按该类ID号），tt3三级信息分类（0不分、其它按该类ID号），c1一级地区分类（0不分、其它按该类ID号），c2二级地区分类（0不分、其它按该类ID号），c3三级地区分类（0不分、其它按该类ID号），xxlx信息类型（0不显1显），xxtj推荐标志（0不显1显），gd固顶标志（0不显1显），tp图片标志（0不显1显），xxsj时间（0不显、1显月日、2显年月日、3显复杂），jj竞价提示（0不显1显），gdxx更多（0不显1显），ys标题颜色（0不显1显），hits点击数（0不显1显），btw标题长度（以字节计），s显示条数，l列数，zt字号，hg行高
            </td>
        </tr>
      </table></td>
    </tr>
	<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="230">
   <tr>
      <td height="28" style="border-bottom-style: solid; border-bottom-width: 1">
<b>调用预览：</b>（效果是按默认的参数。若图片标志不显示请检查网站系统管理－基本参数中的网址是否与当前对应）
</td>
        </tr>

    <tr bgcolor="#BDBEDE">
      <td height="28" style="border-bottom-style: solid; border-bottom-width: 1">
<!--代码开始-->
<script language="javascript" src="../user/newinfo.asp?xxsx=0&tt1=0&tt2=0&tt3=0&c1=0&c2=0&c3=0&xxlx=1&xxtj=0&gd=0&tp=1&xxsj=0&jj=0&gdxx=1&ys=0&hits=0&btw=20&s=10&l=1&zt=13&hg=10"></script>
<!--代码结束-->
</td>
        </tr>
      </table
</table>

