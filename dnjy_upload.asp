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
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/<%=strInstallDir%>css/style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
BODY{
font-size:9pt
}
.tx1 { height: 20px;font-size: 9pt; border: 1px solid; border-color: #000000; color: #0000FF}
-->
</style>
<SCRIPT language=javascript>
function check() 
{
	var strFileName=document.form1.FileName.value;
	if (strFileName=="")
	{
    	alert("请选择要上传的文件");
		document.form1.FileName.focus();
    	return false;
  	}
}
</SCRIPT>

</head>
<BODY BGCOLOR="#ffffff" CLASS="p9">
<FORM METHOD="post" NAME="form1" ACTION="DNJY_upfile.asp?ttup=<%=trim(request("ttup"))%>&frmname=<%=trim(request("frmname"))%>&id=<%=trim(request("id"))%>&oneid=<%=trim(request("oneid"))%>&twoid=<%=trim(request("twoid"))%>&threeid=<%=trim(request("threeid"))%>" ENCTYPE="multipart/form-data">
  <TABLE BORDER="0" ALIGN="center" CELLPADDING="0" CELLSPACING="0">
    <TR> 
      <TD><TABLE  BORDER="1" ALIGN="center" CELLPADDING="0" CELLSPACING="0" BORDERCOLOR="#111111" STYLE="BORDER-COLLAPSE: collapse">
          <input type="hidden" name="filenum" value="1">		  
		    <TR> 
            <TD>
			<input type="file" name="file" size="40">
			<input type="submit" class="button" name="Submit" value="上传图片" ></TD>
          </TR>
        </TABLE></TD>
    </TR>
  </TABLE>
</FORM>
</body>
</html>
