<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../Include/Function.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
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
	Response.Write"<script>alert('�����ɹ���');location='notice.asp'</script>"
	Response.End
End If
%>
<HTML>
 <HEAD>
  <TITLE>֪ͨ��վվ��</TITLE>
<!--kindeditor�༭��-->
		<link rel="stylesheet" href="../KindEditor/themes/default/default.css" />
		<script charset="utf-8" src="../KindEditor/kindeditor-min.js"></script>
		<script charset="utf-8" src="../KindEditor/lang/zh_CN.js"></script>
		<script>
			var editor;			
			KindEditor.ready(function(K) {
				editor = K.create('textarea[name="citynotice"]', {
					resizeType : 2,//0:����kindeditor�༭����С���ɸı�; 1:ֻ�ܸñ�߶�; 2:���Ըı�߶ȺͿ��
					afterBlur: function(){this.sync();},//ʧȥ����ִ��this���ֵ
					allowPreviewEmoticons : false,
					allowImageUpload : false,					
					items : [
						'fontname', 'fontsize', '|', 'forecolor', 'hilitecolor', 'bold', 'italic', 'underline',
						'removeformat', '|', 'justifyleft', 'justifycenter', 'justifyright', 'insertorderedlist',
						'insertunorderedlist', '|', 'emoticons', 'image', 'link']
				});
			});
		</script>
<!--kindeditor�༭��END-->
<script language="javascript">
<!--
function formcheck(){
if (document.form.citynotice.value=="" )
	{
	  alert("����֪ͨ����")	  
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
      <td height="20" bgcolor="#799AE1"><div align="center"><font color="#FFFFFF" style="font-size:14px">ͨ ֪ �� վ վ �� </font></div></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><br />
        <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#D6DFF7">

            <tr>
              <td width="120" bgcolor="#E3EBF9" align="right">֪ͨ���ݣ�</td>
              <td height="250" bgcolor="#E3EBF9">&nbsp;<textarea id="editor" name="citynotice"  style="width:550px;height:280px;visibility:visible"></textarea><br>(��վ�����ڷ�վ�����̨������������)</td>
          </tr>

            <tr>
              <td bgcolor="#E3EBF9">&nbsp;</td>
              <td bgcolor="#E3EBF9">&nbsp;<input type="submit" name="Submit" value="�ύ" /></td>
            </tr>
          </table>
      <br /></td>
    </tr>
  </table>
</form>
 </BODY>
</HTML>