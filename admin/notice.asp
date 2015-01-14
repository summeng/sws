<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../Include/Function.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("02")
Dim Action,citynotice
Call OpenConn
Action=LCase(Request.QueryString("Action"))
if Action="" Then
citynotice=conn.execute ("select citynotice from [DNJY_config]")(0)
Else
	set rs=server.createobject("adodb.recordset")
	sql="select citynotice from [DNJY_config]"
	rs.open sql,conn,1,3
	rs("citynotice")=request("citynotice")
	rs.update
	Call CloseRs(rs)
	Conn.Close:Set Conn=Nothing
	Response.Write"<script>alert('发布成功！');location='notice.asp'</script>"
	Response.End
End If
%>
<HTML>
 <HEAD>
  <TITLE>通知分站站长</TITLE>
<!--kindeditor编辑器-->
		<link rel="stylesheet" href="../KindEditor/themes/default/default.css" />
		<script charset="utf-8" src="../KindEditor/kindeditor-min.js"></script>
		<script charset="utf-8" src="../KindEditor/lang/zh_CN.js"></script>
		<script>
			var editor;			
			KindEditor.ready(function(K) {
				editor = K.create('textarea[name="citynotice"]', {
					resizeType : 2,//0:设置kindeditor编辑器大小不可改变; 1:只能该变高度; 2:可以改变高度和宽度
					afterBlur: function(){this.sync();},//失去焦点执行this获得值
					allowPreviewEmoticons : false,
					allowImageUpload : false,					
					items : [
						'fontname', 'fontsize', '|', 'forecolor', 'hilitecolor', 'bold', 'italic', 'underline',
						'removeformat', '|', 'justifyleft', 'justifycenter', 'justifyright', 'insertorderedlist',
						'insertunorderedlist', '|', 'emoticons', 'image', 'link']
				});
			});
		</script>
<!--kindeditor编辑器END-->
<script language="javascript">
<!--
function formcheck(){
if (document.form.citynotice.value=="" )
	{
	  alert("请填通知内容")	  
	  return false
	 }
}
// --></script>
 </HEAD>
 <BODY>
<form action="?action=save" name="form" method="post" onSubmit="return formcheck()">
<LINK rel="stylesheet" type="text/css" href="../css/style.css">
  <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
    <tr>
      <td height="20" bgcolor="#799AE1"><div align="center"><font color="#FFFFFF" style="font-size:14px">通 知 分 站 站 长 </font></div></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><br />
        <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#D6DFF7">

            <tr>
              <td width="120" bgcolor="#E3EBF9" align="right">通知内容：</td>
              <td height="250" bgcolor="#E3EBF9">&nbsp;<textarea id="editor" name="citynotice"  style="width:550px;height:280px;visibility:visible"></textarea><br>(分站长可在分站管理后台看到以上内容)</td>
          </tr>

            <tr>
              <td bgcolor="#E3EBF9">&nbsp;</td>
              <td bgcolor="#E3EBF9">&nbsp;<input type="submit" name="Submit" value="提交" /></td>
            </tr>
          </table>
      <br /></td>
    </tr>
  </table>
</form>
 </BODY>
</HTML>