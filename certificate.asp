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
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
<script>
function loadThreadFollow(t_id,b_id){
	var targetImg =eval("document.all.followImg" + t_id);
	var targetDiv =eval("document.all.follow" + t_id);
	
	if ("object"==typeof(targetImg)){
		if (targetDiv.style.display!='block'){
			targetDiv.style.display="block";
			
		}else{
			targetDiv.style.display="none";
		}
	}
}
</script>
</head>
<%dim tupian,edit
username=request.cookies("dick")("username")
Call OpenConn
tupian=Request.form("upname")
edit=Request.form("edit")
Dim certificate,certificate1,certificate2,certificate3,certificate4,certificate5,certificate6,certificate7
certificate=conn.Execute("Select certificate From DNJY_config")(0)
certificate=split(certificate,"|")
certificate1=certificate(0)
certificate2=certificate(1)
certificate3=certificate(2)
certificate4=certificate(3)
certificate5=certificate(4)
certificate6=certificate(5)
certificate7=certificate(6)%>

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="980" height="426">
    <tr>
      <td width="214" height="354" valign="top" bgcolor="#FFFFFF"><div align="center">
      
        <!--#include file=userleft.asp--></div></td>
		<td width="5">&nbsp;</td>
<td width="760" height="354" valign="top" bgcolor="#FFFFFF">
<table width="760" height="100%" border="0" align="right" cellpadding="0" cellspacing="0" bordercolor="#111111"  class="ty1">
<tr>
<td width="100%" height="40" align="center" valign="top"><div align="right">
<%if edit="" Then%>
<form method="POST" name="comForm" action="certificate.asp?<%=C_ID%>">
<input TYPE="hidden" name="username" value="<%=username%>">
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse">
        <tr>
         <td height="50" bgcolor="#FAFAFA" align="center" colspan="2"><b>在此提交申请店铺和企业名片审核时需要的证照（证照仅供网站管理员内部审核用，不公开显示）</b></td>                   
         </tr>
  <tr> 
<td align="left" colspan="2">
<p style="margin-left: 60px"><font color=#ff0000><b>站长温馨提示：</b><%=conn.Execute("Select prompt From DNJY_config")(0)%></font></td>
</tr>
<%If certificate1="1" then%>
<tr>
<td height="30" align="right">法人身份证：</td><td><input type="text" name="certificate1" size="47" id="certificate1" maxlength="50" value="<%=rs("certificate1")%>"> <a target="_blank" href="<%=rs("certificate1")%>">[预览]</a></td></tr>
<%End if%>
<%If certificate2="1" then%>
<tr><td height="30" align="right">法人标准像：</td><td><input type="text" name="certificate2" size="47" id="certificate2" maxlength="50" value="<%=rs("certificate2")%>"> <a target="_blank" href="<%=rs("certificate2")%>">[预览]</a></td></tr>
<%End if%>
<%If certificate3="1" then%>
<tr><td height="30" align="right">营业执照：</td><td><input type="text" name="certificate3" size="47" id="certificate2" maxlength="50" value="<%=rs("certificate3")%>"> <a target="_blank" href="<%=rs("certificate3")%>">[预览]</a></td></tr>
<%End if%>
<%If certificate4="1" then%>
<tr><td height="30" align="right">国税登记证：</td><td><input type="text" name="certificate4" size="47" id="certificate2" maxlength="50" value="<%=rs("certificate4")%>"> <a target="_blank" href="<%=rs("certificate4")%>">[预览]</a></td></tr>
<%End if%>
<%If certificate5="1" then%>
<tr><td height="30" align="right">地税登记证：</td><td><input type="text" name="certificate5" size="47" id="certificate2" maxlength="50" value="<%=rs("certificate5")%>"> <a target="_blank" href="<%=rs("certificate5")%>">[预览]</a></td></tr>
<%End if%>
<%If certificate6="1" then%>
<tr><td height="30" align="right">组织机构代码证：</td><td><input type="text" name="certificate6" size="47" id="certificate2" maxlength="50" value="<%=rs("certificate6")%>"> <a target="_blank" href="<%=rs("certificate6")%>">[预览]</a></td></tr>
<%End if%>
<%If certificate7="1" then%>
<tr><td height="30" align="right">其它证照(门面、车间、商标等)：</td><td><input type="text" name="certificate7" size="47" id="certificate2" maxlength="50" value="<%=rs("certificate7")%>"> <a target="_blank" href="<%=rs("certificate7")%>">[预览]</a></td></tr>
<%End if%>
<tr><td height="30" colspan="2">
<span id="followImg1" style="CURSOR: hand" title="选中上传图片" onclick="loadThreadFollow(1,5)"><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;请<font color="#0000FF">上传图片</font>↓（限gif,jpg,bmp,png格式）</span></td>
</tr>                                                  
<tr  style="display:none" id="follow1">
<td height="23" colspan="2"><TABLE width="760">
<p style="margin-left: 60px">
<%If certificate1="1" then%>
<tr><td height="30" align="right">
上传法人身份证</td><td><iframe name="I1" src="DNJY_Upload.asp?ttup=certificate1" scrolling="no" border="0" frameborder="0" width="450" height="35"></td></tr>
</iframe><%End if%>
<%If certificate2="1" then%><tr><td height="30" align="right">
上传法人标准像</td><td>
<iframe name="I1" src="DNJY_Upload.asp?ttup=certificate2" scrolling="no" border="0" frameborder="0" width="450" height="35">
</iframe></td></tr>
<%End if%>
<%If certificate3="1" then%><tr><td height="30" align="right">
上传营业执照</td><td>
<iframe name="I1" src="DNJY_Upload.asp?ttup=certificate3" scrolling="no" border="0" frameborder="0" width="450" height="35">
</iframe></td></tr>
<%End if%>
<%If certificate4="1" then%><tr><td height="30" align="right">
上传国税登记证</td><td>
<iframe name="I1" src="DNJY_Upload.asp?ttup=certificate4" scrolling="no" border="0" frameborder="0" width="450" height="35">
</iframe></td></tr>
<%End if%>
<%If certificate5="1" then%><tr><td height="30" align="right">
上传地税登记证</td><td>
<iframe name="I1" src="DNJY_Upload.asp?ttup=certificate5" scrolling="no" border="0" frameborder="0" width="450" height="35">
</iframe></td></tr>
<%End if%>
<%If certificate6="1" then%><tr><td height="30" align="right">
上传组织机构代码证</td><td>
<iframe name="I1" src="DNJY_Upload.asp?ttup=certificate6" scrolling="no" border="0" frameborder="0" width="450" height="35">
</iframe></td></tr>
<%End if%>
<%If certificate7="1" then%><tr><td height="30" align="right">
上传其它证照(门面、车间、商标等)</td><td>
<iframe name="I1" src="DNJY_Upload.asp?ttup=certificate7" scrolling="no" border="0" frameborder="0" width="450" height="35">
</iframe></td></tr>
<%End if%>
</TABLE>
</td>
</tr>                                            

<tr> 
<td height="38" colspan="2">
<p align="center">
<input type="submit" value="确认保存" name="edit"></td>
</tr>                                      
</table></form>
<%
else 
username=request.cookies("dick")("username")
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,3
If certificate1="1" Then rs("certificate1")=trim(request("certificate1"))
If certificate2="1" Then rs("certificate2")=trim(request("certificate2"))
If certificate3="1" Then rs("certificate3")=trim(request("certificate3"))
If certificate4="1" Then rs("certificate4")=trim(request("certificate4"))
If certificate5="1" Then rs("certificate5")=trim(request("certificate5"))
If certificate6="1" Then rs("certificate6")=trim(request("certificate6"))
If certificate7="1" Then rs("certificate7")=trim(request("certificate7"))
rs.update
response.write "<center>操作成功！</center>"
response.write "<meta http-equiv=refresh content=""1;URL=certificate.asp?"&C_ID&""">"
end if%></td>
</tr>
</table>
<%Call CloseRs(rs)
Call CloseDB(conn)%>
</body>
</html>