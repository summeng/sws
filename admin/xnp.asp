<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
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
    response.Write "<script language=javascript>alert('操作成功！将自动生成配置文件config.asp');location.href='admin_config.asp?page="&url&"';</script>"
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
    <TD height="20" bgcolor="#799AE1" align="center" colspan="13"><FONT color="#FFFFFF" style="font-size:14px">给 VIP 用 户 赠 送 <%=rmb_mc%></FONT></TD>
 </TR>

          <tr> 
            <td bgcolor="#FFFFFF" align="center"> 
              <table width="96%" border="0" cellspacing="0" cellpadding="6">
                <tr> 
                  <td width="10%">赠送数量：</td>
                  <td width="30%"> 
                    <input type="text" name="xnp" value="10" size="20"> 个<%=rmb_mc%>/人
                  <br>（赠送太多没有意义，造成运算时耗费资源，建议1000以下）</td>
                  <td width="60%"> 
                   上次赠送时间：<%=rs("xnpsj")%> &nbsp;&nbsp;&nbsp;&nbsp;赠送数量<%=rs("xnp")%>个<%=rmb_mc%>/人
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
                    <input type="submit" name="Submit" value="确认设置" class="Tips_bo">
                  </div></td>
                  <td><div align="center">
                    <input type="reset" name="Submit" value="取消设置" class="Tips_bo">
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
