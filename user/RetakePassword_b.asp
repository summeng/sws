<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/tt_auto_lock.asp" -->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%response.buffer="True"
dim user,username
dim problem,id
username=request.form("username")
problem=request.form("problem")
dim connuser,password
Call OpenConn
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_user] where username='"&UserName&"'"
rs.open sql,conn,1,1
if rs.eof then
Response.Write "û�д��û�����"
Call CloseRs(rs)
Call CloseDB(conn)
Response.End
End if
session("gpw_error")=session("gpw_error")+1
if rs.eof then response.write "<tr><td colspan=7 bgcolor=#ffffff>&middot;<font color=#ff0000>û�д˻�Ա�����޷��һ����룡���Ѵ�������"&session("gpw_error")&"�Σ����޽�����IP�޷����ʣ�</font></td></tr>"
%>
<meta http-equiv="Content-Language" content="zh-cn">
<link href="/<%=strInstallDir%>css/css.css" rel="stylesheet" type="text/css">
<title>�����һ�</title>
<style>
<!--
.style1 {
	font-size: 16px;
	font-family: "����";
}
-->
</style>

</head>
<% if not rs.eof And rs("problem")<>"" then%>
<body topmargin="0" leftmargin="0">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="400" id="AutoNumber1" height="147">
    <tr> 
      <td width="500" height="25" bgcolor="#FF9900"> <p align="center">&nbsp; ..:::�ڶ�������������Ĵ�:::...</td>
    </tr>
  <tr> 
    <td width="500" height="122"> 
      <table border="0" width="100%" id="table1" height="100">
             <form method="POST" name="form1" onSubmit="return form1_onsubmit()" action="RetakePassword_c.asp">
      <tbody>
        <tr>
          <td width="100%" align="center">��
             �����û�����<%=rs("username")%> </td>
        </tr>
                <tr>
          <td width="100%" align="center">��
            �������⣺<%=rs("problem")%> </td>
        </tr>
       	<tr > 
         <td width="100%" align="center">��
         <span class="style1">����������𰸣�</span><input class="inputa" type="answer1" maxlength="30" name="answer1" size="20" ></td>
                      </tr>      
        
      </tbody>
    </table>
	<input type="hidden" name="username" value="<%=rs("username")%>">
    <center><input type="submit" value=" ����һ�� " name="submit" >
      </td>
		</tr></form>
		</table>
	</td>
  </tr>
  </table>
<%else%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="380" id="AutoNumber1" height="70">
    <tr> 
      <td width="380" height="25" bgcolor="#FF9900"> <p align="center"> </td>
    </tr>
  <tr> 
    <td width="380" > 
   ���������һ����ϲ�ȫ���޷��һ����룬�������Ա��ϵ��   
	</td>
  </tr>
  </table><%end If%>
<table border="0" width="380" id="table2" bgcolor="#FF9900">
	<tr>
		<td>��</td>
	</tr>
</table>
<%Call CloseRs(rs)
Call CloseDB(conn)%>