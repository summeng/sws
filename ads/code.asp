
<!--#include file=cookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
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
  <tr><td height="28" colspan="2"><div align="center" ><font color="#FFFFFF"><b>������ʾ��</b></font></div></td>
 
  </tr>
  <tr>
    <td height="250" align="center" bgcolor="#FFFFFF"> ��JS�ļ�����Ҫʹ�á�����JS�����ܺ�ſ�ʹ�ã�<br>
        <br>		<textarea name="S1" cols="80" rows="6" class="input2"><!--�����뿪ʼ����Ҫ���ھ�̬ҳ��-->
  <script src="ads/JS/���ID��_һ������ID��_��������ID��_��������ID��.js"></script>
  <!--���������-->
      </textarea>

        <br>
		JS���루JS��̬���÷�ʽ��<br>
        <textarea name="S1" cols="80" rows="6" class="input2"><!--�����뿪ʼ�����ڶ�̬ҳ��-->
  <script src="&lt;%=adjs_path("ads/js","���ID��",c1&"_"&c2&"_"&c3)%>"></script>
  <!--���������-->
      </textarea>
        <br>
        <br>
      JS���루ASP���÷�ʽ��<br>
      <br>
      <textarea name="textarea" cols="80" rows="6" class="input2"><!--�����뿪ʼ�����ڶ�̬ҳ��-->
  <script src="ads/ad.asp?adid=���ID��&c=&lt;%=c1%>,&lt;%=c2%>,&lt;%=c3%>"></script>
<!--���������-->
      </textarea></td> 
  </tr>
  <tr>
    <td height="80" bgcolor="#FFFFFF"><table width="98%"  border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td>���������ġ����ID�š����ĳ���Ӧ����ID��<br>
		������ĵ���ID�Ÿ��ĳ���Ӧ������ID��<br>
*�Ƽ�! ��JS�ļ����룬����Ҫ������ִ�У��ٶȵ�һ��ȱ�㲻ͳ����ʾ������<br>
*������ͨ�̶���ʾ��ͼƬ/FLASH��棬������Ĵ��븴�Ƶ�Ҫ��ʾ����λ�á�<br>
*�����������͹�渴�Ƶ���ʾҳ�������λ�ü��ɡ�</td>
      </tr>
    </table></td>
  </tr>
</table>
<br>
</body>
</html>