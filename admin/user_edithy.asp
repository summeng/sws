<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Call OpenConn
dim username,id,vip,rs5,sql5,vipys,dqts
id=trim(request("id"))
if not isnumeric(id) or id="" then
response.write "<li>��������"
response.end
end if
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select sdkg,sdname,sdjj,vip,Amount,dpdata,sdfg from [DNJY_user] where id="&cstr(id) 
rs.open sql,conn,1,1
%>
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<title>�޸Ļ�Ա����</title>
<script language="JavaScript">
function CheckForm()
{
	
	
		if (document.alipayment.vipys.value.length == 0) {
		alert("������VIP���������򱣳�Ϊ0.");
		document.alipayment.vipys.focus();
		return false;
	}
                if (isNaN(document.alipayment.vipys.value))
	{
	  alert("VIP��������ӦΪ���֣�")
	  document.alipayment.vipys.focus()
	  return false
	 }

	return true;
}

</script>
<body topmargin="0" leftmargin="0">
<p>��</p>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="400" height="100">
    <form action="?hyjb=yes&id=<%=id%>" name="alipayment" method="POST" onSubmit="return CheckForm();">
    <tr>
      <td width="130" height="23">
      <p align="center">��Ա����</td>
      <td width="270" height="23"><select id="ctlvip" name="vip">
  <option value="0" <%if rs("vip")=0 then%>selected<%end if%>>��ͨ��Ա</option>
  <option value="1" <%if rs("vip")=1 then%>selected<%end if%>>VIP��Ա</option>
</select> </td>
    </tr>
                        <TR class=tdbg>
                          <TD width="16%" >����������</TD>
                          <TD width="86%" ><input onKeyPress="return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,/^\d*\.?\d{0,2}$/,&#9;event.dataTransfer.getData('Text'))" maxLength=10 size=10 name="vipys" value="<%=vipsj%>" class="form">
                            ���£�ָ����VIP��Ա�ʸ������,���¼ƣ�<br>���û��ʽ����<%=rs("Amount")%>Ԫ����������<%=Fix(rs("Amount")/vipje)%>���»�Ա�ʸ�<br>��������12���£��Ա��ڹ���</TD>
                        </TR>             
    <tr>
      <td width="130" height="12"></td>
      <td width="270" height="12">
      
      <input type="submit" value="�ύ" name="B1" style="color: #FFFFFF; border: 1px solid #000000; background-color: #666666"></td>
    </tr>
    </form>
  </table>
  </center>
</div>
<SCRIPT>
	function regInput(obj, reg, inputStr)
	{
		var docSel	= document.selection.createRange()
		if (docSel.parentElement().tagName != "INPUT")	return false
		oSel = docSel.duplicate()
		oSel.text = ""
		var srcRange	= obj.createTextRange()
		oSel.setEndPoint("StartToStart", srcRange)
		var str = oSel.text + inputStr + srcRange.text.substr(oSel.text.length)
		return reg.test(str)
	}
</SCRIPT>

<%If request("hyjb")="yes" Then
id=request("id")
set rs5=server.createobject("adodb.recordset")
sql5="select username,name,vipdq,vip,Amount from [DNJY_user] where id="&id
rs5.open sql5,conn,1,3
if rs5.eof or rs5.bof then
response.write "<li>��������"
response.end
end If

if CInt(request("vipys"))<=0 And request("vip")=1 then
Response.Write ("<script language=javascript>alert('����VIP�ʸ������������0!');history.go(-1);</script>")
response.end
end If
if  rs5("Amount")<0 Or rs5("Amount")<vipje*CInt(request("vipys")) And request("vip")=1 then
Response.Write ("<script language=javascript>alert('���û����ʻ����㣬���ֵ��������VIP��Ա�ʸ�!');history.go(-1);</script>")
response.end
end If
rs5("vip")=request("vip")
if request("vip")=1 Then
vipys=CInt(request("vipys"))
dqts=vipys*30
dqts= dateadd("d", dqts, now)
rs5("vipdq")=dqts
end if
rs5.update
if request("vip")=1 Then
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_Financial] "
rs.open sql,conn,1,3
rs.addnew
rs("username")=rs5("username")
rs("aliname")=rs5("name")
rs("ywje")=vipje*vipys
rs("ywlx")="֧��"
rs("czbz")="����VIP��Ա�ʸ�"
rs("czz")=request.cookies("administrator")
rs("admincl")=1
rs("addTime")=now()
rs.update
Call CloseRs(rs)
conn.execute "UPDATE DNJY_user SET Amount=Amount-"&vipje*vipys&" WHERE username='"&rs5("username")&"'" 'ͬʱ���û������ʻ����
conn.execute "UPDATE DNJY_user SET tAmount=tAmount+"&vipje*vipys&" WHERE username='"&rs5("username")&"'" 'ͬʱ���û����������ܽ��
End if
rs5.close
set rs5=Nothing
Response.Write ("<script language=javascript>alert('�����ɹ�!');history.go(-1);</script>")
End If
Call CloseDB(conn)%>