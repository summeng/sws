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
dim action,bigclassname,founderr,errmsg,bigclassid,url,hide,rs2,closeconn
action=trim(request("action"))
bigclassid=trim(request("bigclassid"))
bigclassname=trim(request("bigclassname"))
url=trim(request("url"))
hide=request("hide")
if action="add" then
		dim classid
		set rs2 = server.CreateObject ("adodb.recordset")
		sql="select bigclassid from DNJY_bigClass where bigclassid="+cstr(request("bigclassid"))
 		rs2.open sql,conn,1,1
		if not rs2.eof and not rs2.bof then
		rs2.close
		set rs2=nothing%>
		<script language=javascript>  
			alert( "��������ID�����ظ�"  );
			location.href = "javascript:history.back()"  
		</script>
		<%
		response.End()
		end if
	if bigclassname="" then
		founderr=true
		errmsg=errmsg & "<br><li>���´���������Ϊ�գ�</li>"
	end if
	if bigclassid="" then
		founderr=true
		errmsg=errmsg & "<br><li>���´���������Ϊ�գ�</li>"
	end if
	if founderr<>true then
		set rs=server.createobject("adodb.recordset")
		rs.open "select * from DNJY_bigClass where bigclassname='" & bigclassname & "'",conn,1,3
		if not (rs.bof and rs.eof) then
			founderr=true
			errmsg=errmsg & "<br><li>���´��ࡰ" & bigclassname & "���Ѿ����ڣ�</li>"
		else
    	 	rs.addnew
     		rs("bigclassid")=bigclassid
			rs("bigclassname")=bigclassname
			if hide="yse" then
			rs("hide")=true
			else
			rs("hide")=false
			end if
			rs("url")=url
			if url<>"" then
			rs("urlok")=true
			end If
			rs("userdraft")=trim(request("userdraft"))			
    	 	rs.update
     		rs.close
	     	set rs=nothing
    	 	'call closeconn()
			response.redirect "classmanage.asp"  
		end if
	end if
end if
if founderr=true then
	response.Write ""&errmsg&""
else
%>
<script language="javascript" type="text/javascript">
function checkbig()
{
   if (document.form1.bigclassname.value=="")
  {
    alert("�������Ʋ���Ϊ�գ�");
    document.form1.bigclassname.focus();
    return false;
  }
    if (document.form1.bigclassid.value=="")
  {
    alert("����������Ϊ�գ�");
    document.form1.bigclassid.focus();
    return false;
  }
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
<BODY>
<script language=javascript src=../Include/mouse_on_title.js></script>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> <br>
      <a href="classmanage.asp"><strong> �� �� �� ��</strong></a> <br>
      <form name="form1" method="post" action="classaddbig.asp" onsubmit="return checkbig()">
        <table width="90%" border="0" align="center" cellpadding="0" cellspacing="2" class="border">
          <tr bgcolor="#999999" class="title"> 
            <td height="25" colspan="2" align="center"><strong>��Ӵ���</strong></td>
          </tr>
          <tr bgcolor="#e3e3e3" class="tdbg"> 
            <td width="126" height="22" bgcolor="#c0c0c0"> <div align="right"><strong>�������ƣ�</strong></div></td>
            <td width="218" bgcolor="#e3e3e3"> <input name="bigclassname" type="text" size="20" maxlength="30"> 
            </td>
          </tr>
          <tr bgcolor="#e3e3e3" class="tdbg"> 
            <td width="126" height="22" bgcolor="#c0c0c0"> <div align="right"><strong>��������</strong></div></td>
            <td width="218" bgcolor="#e3e3e3"> <input name="bigclassid" type="text" size="20" maxlength="30"  onKeyUp="if(isNaN(value)){alert('ֻ������������');value='';}"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))"> 
            </td>
          </tr>
		  <tr bgcolor="#e3e3e3" class="tdbg"> 
            <td width="126" height="22" bgcolor="#c0c0c0"> <div align="right"><strong>�����ӵ�ַ��</strong></div></td>
            <td bgcolor="#e3e3e3"> <input name="url" type="text" id="url" size="20" maxlength="50"> 
            (�����д��ַֻ���ⲿ����,�����ڲ�������Ѷ��Ŀ! )           </td>
          </tr>
		  <tr bgcolor="#e3e3e3" class="tdbg"> 
            <td width="126" height="22" bgcolor="#c0c0c0"> <div align="right"><strong>��ҳ��ʾ��</strong></div></td>
            <td width="218" bgcolor="#e3e3e3"><label>
              <input name="hide" type="checkbox" id="hide" value="yes" checked="checked" />
            ��ָ��ҳ��飩</label></td>
          </tr>
		  <tr bgcolor="#e3e3e3" class="tdbg"> 
            <td width="126" height="22" bgcolor="#c0c0c0"> <div align="right"><strong>�Ƿ������ԱͶ�壺</strong></div></td>
            <td width="218" bgcolor="#e3e3e3"><label><input name="userdraft" type="radio" value="1" checked>��
           <input name="userdraft" type="radio" value="0">��</label></td>
          </tr>
           <tr bgcolor="#c0c0c0" class="tdbg"> 
            <td height="22" align="center" bgcolor="#c0c0c0">&nbsp; </td>
            <td height="22" align="center" bgcolor="#e3e3e3"> <div align="left"> 
                <input name="action" type="hidden" id="action" value="add">
                <input name="add" type="submit" value=" �� �� ">
              </div></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
<%end if
Call CloseDB(conn)%>