
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
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>������Ϣ�ٱ�</title>
<link rel="stylesheet" href="/<%=strInstallDir%>css/style.css">
<script LANGUAGE=JavaScript>
  function textLimitCheck(thisArea, maxLength){
    if (thisArea.value.length > maxLength){
      alert(maxLength + ' ��������. \r�����Ľ��Զ�ȥ��.');
      thisArea.value = thisArea.value.substring(0, maxLength);
      thisArea.focus();
    }
    /*��дspan��ֵ����ǰ��д���ֵ�����*/
    messageCount.innerText = thisArea.value.length;
  }
</script>

<script language="javascript">
<!--
//��֤������ȷ��
function checkemail(Bad_infomail){
var str=Bad_infomail;
//��JavaScript�У�������ʽֻ��ʹ��"/"��ͷ�ͽ���������ʹ��˫����
var Expression=/\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*/; 
var objExp=new RegExp(Expression);
if(objExp.test(str)==true){
return true;
}else{
return false;
}
}


/*****  ��ѡ��ť��� �˶β����޸� ���������*****/
function check_radio(check_radio)
{
    for(i=0;i<check_radio.length;i++){
            if(check_radio[i].checked==true){
                return true;
            }
        }    
    return false;
}
/*****  ��ѡ��ť���end  *****/

//���ж�
function chkinput(){

//����7��Ϊ��ѡ��ť��⣬���������
var form1 = document.getElementById("form");
if (!check_radio(form.Bad_infolx))
{
alert("����ȷѡ��ٱ����ͣ�");
form.Bad_infolx[0].focus();
return false;
}

  if ((!checkemail(form.Bad_infomail.value)) && (document.form.Bad_infomail.value!="" ))
	{
    alert("������Email��ַ����ȷ������������!");
    document.form.Bad_infomail.focus();
    return false;
    }
if (document.form.validate_codee.value==0)
	{
	  alert("������֤��")
	  document.form.validate_codee.focus()
	  return false
	 }
}
// --></script>
</head>
<body>
<TABLE width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
<TR>
	<TD align="center"><b><font color="#FF0000">������ʵ���͹۶Դ�</font></b></TD>
</TR>
</TABLE>
<TABLE width="80%" border="0" align="center" cellpadding="0" cellspacing="1">
<TR>
	<TD><b><font color="#FF0000"><b>�ٱ�����(��ѡ)��</b><p></TD></TR>
<form name=form action=Bad_info_chk.asp  method=post onsubmit="return chkinput()">
<TR>
	<TD>
<input type="radio" name="Bad_infolx" value="��Ϣ���������"> ��Ϣ���������<br>
<input type="radio" name="Bad_infolx" value="��Ϣ����ȨΥ����"> ��Ϣ����ȨΥ���ַ�<br>
<input type="radio" name="Bad_infolx" value="������ŵ��������"> ������ŵ��������<br>
<input type="radio" name="Bad_infolx" value="ƭȡ����Ǯ��"> ƭȡ����Ǯ��<br> 
<input type="radio" name="Bad_infolx" value="��Ϣ�漰��Ʒ��ðα��"> ��Ϣ�漰��Ʒ��ðα��<br> 
<input type="radio" name="Bad_infolx" value="��Ϣ�漰��ƷΪΥ����Ʒ"> ��Ϣ�漰��ƷΪΥ����Ʒ<br>
<input type="radio" name="Bad_infolx" value="������Ϣ����"> ������Ϣ����<br>
<input type="radio" name="Bad_infolx" value="����ԭ��"> ����ԭ��<p> 

<b>����˵��(ѡ��)��</b><p>
�뾡����ϸ��ӳʵ�����,�Ա����Ǽ�ʱ��ȷ�Ĵ������ǵľٱ�<br>

<textarea name="Bad_info"  cols="50" rows="5" class="input2" onkeyUp="textLimitCheck(this, 200);"></textarea><br>
�� 200 ���ַ� ������ <font color="#CC0000"><span id="messageCount">0</span></font> ����<p>

<b>��������(ѡ��)��</b> <input type="text" name="Bad_infomail" value="" size="30" maxLength=50><p>
��ʽ��֤��<img src="../tt_getcodee.asp" alt="��֤��,�������?����ˢ����֤��" height="10" style="cursor : pointer;" onClick="this.src='../tt_getcodee.asp?t='+(new Date().getTime());" /><input type="text" name="validate_codee" size="4" onkeyup="value=value.replace(/[^\d|^\.]/g,'')"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))">&nbsp;&nbsp; <img src=../images/memo.gif alt="��֤��,�������?����ˢ����֤��">
<input type="hidden" name="infoid" value="<%=Request.QueryString("id")%>">
<input type="hidden" name="city_oneid" value="<%=request("city_oneid")%>">
<input type="hidden" name="city_twoid" value="<%=request("city_twoid")%>">
<input type="hidden" name="city_threeid" value="<%=request("city_threeid")%>">
</TD>
</TR>
</TABLE>
	  <TABLE width="80%" border="0" align="center" cellpadding="0" cellspacing="1">
          <tr>
            <td height="35"><div align="center">
              <input name="AddAD" type="submit" class="input1" value="ȷ���ٱ�">
            </div></td>
            <td height="35"><div align="center">
              <input name="AddAD" type="reset" class="input1" value="ȡ���ٱ�">
            </div></td>
          </tr>
        </table>
  </form>
<TABLE width="80%" border="0" align="center" cellpadding="0" cellspacing="1">
             <tr>
                <td height="6" bgcolor="#FBC831">
��ÿ���ͬһ��Ϣֻ�ܾٱ�һ��<br>
���ǻᾡ��������ľٱ�����ϢԴ���к�ʵ,лл���Ա�վ��֧��!<br>
��վ�Ծٱ�����Ϣֻ��һ�㴦�����Դ�������Ϣ��ʵ�ԺͺϷ��Ը���</font></td>
                </tr>
</TABLE>

</body>
</html>