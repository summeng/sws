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
<%call checkmanage("03")
Call OpenConn
	dim url 
	xnp=request("xnp")
	set rs=server.createobject("adodb.recordset")
	sql="select * from [DNJY_config]"
	rs.open sql,conn,1,3
    if request("ok")=1 then
    rs("xnp")=request("xnp")
    rs("xnpsj")=now()
    Conn.Execute("Update DNJY_user Set hb=hb+"&xnp&" where vip=1")
    rs.update
	url="xnp.asp"
    response.Write "<script language=javascript>alert('�����ɹ������Զ����������ļ�config.asp');location.href='admin_config.asp?page="&url&"';</script>"
    response.end
    else
%>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
}
.style1 {
	color: #FFFFFF;
	font-weight: bold;
	font-size: 12px;
}
-->
</style><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center">
      <form name="form1" method="POST" action="xnp.asp?ok=1">
        <table border="1" cellspacing="0" cellpadding="0" bgcolor="#F5F5F5" width="100%" style="border-collapse: collapse" bordercolor="#ADAED6">
	<TR>
    <TD height="20" bgcolor="#799AE1" align="center" colspan="13"><FONT color="#FFFFFF" style="font-size:14px">�� VIP �� �� �� �� <%=rmb_mc%></FONT></TD>
 </TR>

          <tr> 
            <td bgcolor="#FFFFFF" align="center"> 
              <table width="96%" border="0" cellspacing="0" cellpadding="6">
                <tr> 
                  <td width="10%">����������</td>
                  <td width="30%"> 
                    <input type="text" name="xnp" value="10" size="20"> ��<%=rmb_mc%>/��
                  <br>������̫��û�����壬�������ʱ�ķ���Դ������1000���£�</td>
                  <td width="60%"> 
                   �ϴ�����ʱ�䣺<%=rs("xnpsj")%> &nbsp;&nbsp;&nbsp;&nbsp;��������<%=rs("xnp")%>��<%=rmb_mc%>/��
                  </td>
                </tr>

              </table>
            </td>
          </tr>
          <tr>
            <td height="35" align="center" bgcolor="#eeeeee"> 
              <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td><div align="center">
                    <input type="submit" name="Submit" value="ȷ������" class="Tips_bo">
                  </div></td>
                  <td><div align="center">
                    <input type="reset" name="Submit" value="ȡ������" class="Tips_bo">
                  </div></td>
                </tr>
              </table>
            </td>
          </tr>

        </table>
<%
    end if
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
%>
      </form>
    </td>
  </tr>
</table>
