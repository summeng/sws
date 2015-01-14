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
<%dim tupian,edit,m,rsm
username=request.cookies("dick")("username")
Call OpenConn
set rsm=conn.execute("select count(id) from [DNJY_info] where yz=1 and username='"&username&"'")
m=rsm(0)
rsm.close
tupian=Request.form("upname")
edit=Request.form("edit")%>

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="980" height="426">
    <tr>
      <td width="214" height="354" valign="top" bgcolor="#FFFFFF"><div align="center">
      
        <!--#include file=userleft.asp--></div></td>
		<td width="5">&nbsp;</td>
<td width="760" height="354" valign="top" bgcolor="#FFFFFF"><table width="760" height="100%" border="0" align="right" cellpadding="0" cellspacing="0" bordercolor="#111111"  class="ty1">
<tr>
          <td width="100%" height="362" align="center" valign="top"><div align="right">

<%if m<xinxis then
response.write "<br>很抱歉，您发布的信息只有 "&m&" 条，要发布 "&xinxis&" 条以上有价值并经审核通过的信息才能修改店铺或企业名片，继续努力噢！"
else
if edit="" Then
%>

<form method="POST" name="comForm" action="comgg.asp?<%=C_ID%>">
<input TYPE="hidden" name="username" value="<%=username%>">
              <table width="100%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse">
        <tr>
                    <td height="25" bgcolor="#FAFAFA"><span class="style5">在此修改店铺企业的标志和公告</span></td>
                    <td bgcolor="#FAFAFA">&nbsp;</td>
                  </tr>
                                                     
<tr>
<td height="30" align="left">
<p style="margin-left: 60px">店铺LOGO地址：&nbsp;&nbsp;<input type="text" name="tupian" size="47" id="tupian" maxlength="200" value="<%=rs("logo")%>"> <%If vip="1" then%><br>店铺Banner地址：<input type="text" name="sdBanner" size="47" id="sdBanner" maxlength="200" value="<%=rs("sdBanner")%>"><%End if%><span id="followImg1" style="CURSOR: hand" title="选中上传店企标图片" onclick="loadThreadFollow(1,5)"><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;无现成图片,可<font color="#0000FF">上传图片</font>↓（限gif,jpg,bmp,png格式）</span></td>
</tr>                                                      
<tr  style="display:none" id="follow1">
<td height="23" align="left">
<p style="margin-left: 60px">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
以下文件1为店铺LOGO，宽高为220*180<iframe name="I1" src="DNJY_Upload.asp?ttup=sdlogo" scrolling="no" border="0" frameborder="0" width="450" height="35">
</iframe><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>
<%If vip="1" then%>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
以下文件1为店铺Banner，宽高为960*150
<iframe name="I1" src="DNJY_Upload.asp?ttup=sdBanner" scrolling="no" border="0" frameborder="0" width="450" height="35">
</iframe>
<%End if%>
</td>
</tr>
                                            
<tr> 
<td height="27" width="611" align="left">
<p style="margin-left: 60px">企业/店铺公告:<textarea rows="9" name="comgg" cols="40"><%=rs("comgg")%></textarea></td>
</tr>
<tr> 
<td height="38">
<p align="center">
<input type="submit" value="确认信息，提交修改" name="edit"></td>
</tr>                                      
</table></form>
<%
else 
username=request.cookies("dick")("username")
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,3
rs("logo")=trim(request("tupian"))
rs("sdBanner")=trim(request("sdBanner"))
rs("comgg")=trim(request("comgg"))
rs("sdkg")=2
rs.update
response.write "<center>店企信息修改成功，要经重新审核才能显示！</center>"
response.write "<meta http-equiv=refresh content=""1;URL=comgg.asp?"&C_ID&""">"
end if%></td>
</tr>
</table>
<%end If
Call CloseRs(rs)
Call CloseDB(conn)%>
</body>
</html>