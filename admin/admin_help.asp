<!--#include file="../Conn.asp" -->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<%call checkmanage("02")%>
<%Dim Action,Data
Action=LCase(Request.QueryString("Action"))
if Action="" Then
	Data=Split(Replace(Application(CacheName&"_help"),"<br>",vbcrlf),"�ӡӡ�")
Else

	Data=Request.Form("help1")&"�ӡӡ�"&Request.Form("help2")&"�ӡӡ�"&Request.Form("help3")&"�ӡӡ�"&Request.Form("help4")&"�ӡӡ�"&Request.Form("help5")&"�ӡӡ�"&Request.Form("help6")



	Conn.Execute("UPdate [DNJY_config] Set help='"&Data&"'")
	
	Application.Lock
	Application(CacheName&"_help")=Data
	Application.unLock
	Conn.Close:Set Conn=Nothing
	Call Alert ("�޸ĳɹ�","admin_help.asp")	
	Response.End()
End If
%>
<HTML>
 <HEAD>
  <TITLE>��վ�����ۺϹ���</TITLE>
  <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
  <LINK rel="stylesheet" type="text/css" href="../css/style.css">
<!--kindeditor�༭��-->

		<link rel="stylesheet" href="../KindEditor/themes/default/default.css" />
		<script charset="utf-8" src="../KindEditor/kindeditor.js"></script>
		<script charset="utf-8" src="../KindEditor/lang/zh_CN.js"></script>
		<script>
			KindEditor.ready(function(K) {
				K.create('#editor1', {
					urlType : 'relative',
					afterBlur: function(){this.sync();},//ʧȥ����ִ��this���ֵ					
					uploadJson : '../KindEditor/asp/upload_json.asp?action=help'
				});
				K.create('#editor2', {
					urlType : 'absolute',
					afterBlur: function(){this.sync();},//ʧȥ����ִ��this���ֵ					
					uploadJson : '../KindEditor/asp/upload_json.asp?action=help'
				});
				K.create('#editor3', {
					urlType : 'domain',
					afterBlur: function(){this.sync();},//ʧȥ����ִ��this���ֵ					
					uploadJson : '../KindEditor/asp/upload_json.asp?action=help'
				});
				K.create('#editor4', {
					urlType : 'domain',
					afterBlur: function(){this.sync();},//ʧȥ����ִ��this���ֵ					
					uploadJson : '../KindEditor/asp/upload_json.asp?action=help'
				});
				K.create('#editor5', {
					urlType : 'domain',
					afterBlur: function(){this.sync();},//ʧȥ����ִ��this���ֵ					
					uploadJson : '../KindEditor/asp/upload_json.asp?action=help'
				});
				K.create('#editor6', {
					urlType : 'domain',
					afterBlur: function(){this.sync();},//ʧȥ����ִ��this���ֵ					
					uploadJson : '../KindEditor/asp/upload_json.asp?action=help'
				});
			});
		</script>
<!--kindeditor�༭��END-->
<script language = "JavaScript">
//˵���ֽ������ƣ����������=========
String.prototype.Length=function(){
temp=this.replace(/([^\x00-\xff])/g,"$1$1");
return temp.length
}
function formcheck(){
var editor1=KindEditor.create('#editor1');
if (editor1.isEmpty())
	{
	  alert("���ڱ�վ���ݲ���Ϊ�գ�")	  
	  return false
	 }
var editor2=KindEditor.create('#editor2');
if (editor2.isEmpty())
	{
	  alert("��Ա�������ݲ���Ϊ�գ�")	  
	  return false
	 } 
var editor3=KindEditor.create('#editor3');
if (editor3.isEmpty())
	{
	  alert("�����������ݲ���Ϊ�գ�")	  
	  return false
	 } 
var editor4=KindEditor.create('#editor4');
if (editor4.isEmpty())
	{
	  alert("�����������ݲ���Ϊ�գ�")	  
	  return false
	 } 
var editor5=KindEditor.create('#editor5');
if (editor5.isEmpty())
	{
	  alert("���������ݲ���Ϊ�գ�")	  
	  return false
	 } 
var editor6=KindEditor.create('#editor6');
if (editor6.isEmpty())
	{
	  alert("��ϵ�������ݲ���Ϊ�գ�")	  
	  return false
	 } 
}
</script>
 </HEAD>
  <BODY>
<form action="?action=save" name="myform" method="post" onSubmit="return formcheck()">
  <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
    <tr>
      <td height="20" bgcolor="#799AE1"><div align="center"><font color="#FFFFFF" style="font-size:14px">�� վ �� �� �� �� �� �� </font></div></td>
    </tr>
    <tr> 
      <td bgcolor="#FFFFFF"><br />
        <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#D6DFF7">

             <tr>
              <td width="100" bgcolor="#E3EBF9" align="right">���ڱ�վ��</td>
              <td height="250" bgcolor="#E3EBF9">&nbsp;<textarea id="editor1" name="help1" style="width:670px;height:400px;"><%=HTMLDecode(Data(0))%></textarea></td>
          </tr>
            <tr>
              <td bgcolor="#E3EBF9" align="right">��Ա������</td>
              <td height="250" bgcolor="#E3EBF9">&nbsp;<textarea id="editor2" name="help2" style="width:670px;height:400px;"><%=HTMLDecode(Data(1))%></textarea></td>
            </tr>
            <tr>
              <td bgcolor="#E3EBF9" align="right">�������</td>
              <td height="250" bgcolor="#E3EBF9">&nbsp;<textarea id="editor3" name="help3" style="width:670px;height:400px;"><%=HTMLDecode(Data(2))%></textarea></td>
            </tr>
            <tr>
              <td bgcolor="#E3EBF9" align="right">����������</td>
              <td height="250" bgcolor="#E3EBF9">&nbsp;<textarea id="editor4" name="help4" style="width:670px;height:400px;"><%=HTMLDecode(Data(3))%></textarea></td>
            </tr>
            <tr>
              <td bgcolor="#E3EBF9" align="right">������</td>
              <td height="250" bgcolor="#E3EBF9">&nbsp;<textarea id="editor5" name="help5" style="width:670px;height:400px;"><%=HTMLDecode(Data(4))%></textarea></td>
            </tr>
            <tr>
              <td bgcolor="#E3EBF9" align="right">��ϵ���ǣ�</td>
              <td height="250" bgcolor="#E3EBF9">&nbsp;<textarea id="editor6" name="help6" style="width:670px;height:400px;"><%=HTMLDecode(Data(5))%></textarea></td>
            </tr>
            <tr>
              <td bgcolor="#E3EBF9" align="right"></td>
              <td height="30" bgcolor="#E3EBF9">&nbsp;
			<input type="checkbox" name="T_SaveImg" value="1" /> ����������Զ��ͼƬ������&nbsp;&nbsp;<input type="checkbox" name="AspJpeg" value="1" checked/>���ϴ�ͼƬ��ˮӡ</td>
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