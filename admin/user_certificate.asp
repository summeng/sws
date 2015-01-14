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
<%Call OpenConn
dim id
id=trim(request("id"))
if not isnumeric(id) or id="" then
response.write "<li>参数错误！"
response.end
end if
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select username,certificate1,certificate2,certificate3,certificate4,certificate5,certificate6,certificate7 from [DNJY_user] where id="&cstr(id)
rs.open sql,conn,1,1

Dim certificate,certificate1,certificate2,certificate3,certificate4,certificate5,certificate6,certificate7
certificate=conn.Execute("Select certificate From DNJY_config")(0)
certificate=split(certificate,"|")
certificate1=certificate(0)
certificate2=certificate(1)
certificate3=certificate(2)
certificate4=certificate(3)
certificate5=certificate(4)
certificate6=certificate(5)
certificate7=certificate(6)%>
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<title>会员申请店企证照</title>
<body topmargin="0" leftmargin="0">
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" >    
    <tr>
      <td height="50" align="center" colspan="2"><b>会员<%=rs("username")%>提交的证照</b></td>
      </td>
    </tr>
	<%If certificate1="1" then%>
    <tr>
      <td align="right">法人身份证：</td>
      <td><%if IsNull(rs("certificate1")) Then
      response.write "尚未提供。"
	  Else
	  response.write"<img src=""../"&rs("certificate1")&"""  BORDER=""0""/>"
	  end if%></td>
    </tr>
	<%End if%>
	<%If certificate2="1" then%>
    <tr>
      <td align="right">法人标准像：</td>
      <td><%if IsNull(rs("certificate2")) Then
      response.write "尚未提供。"
	  Else
	  response.write"<img src=""../"&rs("certificate2")&"""  BORDER=""0""/>"
	  end if%></td>
    </tr>
	<%End if%>
	<%If certificate3="1" then%>
    <tr>
      <td align="right">营业执照：</td>
      <td><%if IsNull(rs("certificate3")) Then
      response.write "尚未提供。"
	  Else
	  response.write"<img src=""../"&rs("certificate3")&"""  BORDER=""0""/>"
	  end if%></td>
    </tr>
	<%End if%>
	<%If certificate4="1" then%>
    <tr>
      <td align="right">国税登记证：</td>
      <td><%if IsNull(rs("certificate4")) Then
      response.write "尚未提供。"
	  Else
	  response.write"<img src=""../"&rs("certificate4")&"""  BORDER=""0""/>"
	  end if%></td>
    </tr>
	<%End if%>
	<%If certificate5="1" then%>
    <tr>
      <td align="right">地税登记证：</td>
      <td><%if IsNull(rs("certificate5")) Then
      response.write "尚未提供。"
	  Else
	  response.write"<img src=""../"&rs("certificate5")&"""  BORDER=""0""/>"
	  end if%></td>
    </tr>
	<%End if%>
	<%If certificate6="1" then%>
    <tr>
      <td align="right">组织机构代码证：</td>
      <td><%if IsNull(rs("certificate6")) Then
      response.write "尚未提供。"
	  Else
	  response.write"<img src=""../"&rs("certificate6")&"""  BORDER=""0""/>"
	  end if%></td>
    </tr>
	<%End if%>
	<%If certificate7="1" then%>
    <tr>
      <td align="right">其它证照(门面、车间、商标等)：</td>
      <td><%if IsNull(rs("certificate7")) Then
      response.write "尚未提供。"
	  Else
	  response.write"<img src=""../"&rs("certificate7")&"""  BORDER=""0""/>"
	  end if%></td>
    </tr>
	<%End if%>
  </table>
  </center>
</div>
<%
Call CloseRs(rs)
Call CloseDB(conn)
%>
