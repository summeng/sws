<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=top.asp-->
<!--#include file=usercookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">


<body topmargin="0" leftmargin="0">

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" height="371">
    <tr>
      <td width="202" height="299" valign="top" background="img/bg.gif"><div align="center"><!--#include file=userleft.asp-->
      </div>
        </td>
      <td width="12" height="299" background="img/line_01.gif">　</td>
      <td width="568" height="299" valign="top" align="center">
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="109">
<%
dim jf,a,b,ct,d,hb,bbb,g_s,aliname,Amount2,czbz,hb_h,hb_j
username=request.cookies("dick")("username")

sub dick_1()
i=1
a=trim(request.form("a"))
b=trim(request.form("b"))
ct=trim(request.form("ct"))
d=trim(request.form("d"))
if a=0 and b=0 And ct=0 And d=0 then
response.write "<script language=JavaScript>" & chr(13) & "alert('请选择转换数量！');" &"window.location='props.asp?"&C_ID&"'" & "</script>"
response.end
end if
set rs=server.createobject("adodb.recordset")
sql = "select a,b,c,d,hb from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,3
if rs("a")<int(request.form("a")) then
response.write "<script language=JavaScript>" & chr(13) & "alert('参数错误！');" &"window.location='props.asp?"&C_ID&"'" & "</script>"
response.end
end if
if rs("b")<int(request.form("b")) then
response.write "<li>参数错误！"
response.write "<script language=JavaScript>" & chr(13) & "alert('参数错误！');" &"window.location='props.asp?"&C_ID&"'" & "</script>"
response.end
end if
if rs("c")<int(request.form("ct")) then
response.write "<script language=JavaScript>" & chr(13) & "alert('参数错误！');" &"window.location='props.asp?"&C_ID&"'" & "</script>"
response.end
end if
if rs("d")<int(request.form("d")) then
response.write "<script language=JavaScript>" & chr(13) & "alert('参数错误！');" &"window.location='props.asp?"&C_ID&"'" & "</script>"
response.end
end if
hb=a*g_a+b*g_b+ct*g_c+d*g_d
rs("a")=rs("a")-int(a)
rs("b")=rs("b")-int(b)
rs("c")=rs("c")-int(ct)
rs("d")=rs("d")-int(d)
rs("hb")=rs("hb")+int(hb)
rs.update
Call CloseRs(rs)
response.write "<script language=JavaScript>" & chr(13) & "alert('转换成功！');" &"window.location='props.asp?"&C_ID&"'" & "</script>"
end sub

sub dick_2()
hb_h=trim(request.form("hb_h"))
set rs=server.createobject("adodb.recordset")
sql = "select jf,hb from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,3
if hb_h="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('请选择转换数量！');" &"window.location='props.asp?"&C_ID&"'" & "</script>"
response.end
end if
rs("jf")=rs("jf")-int(hb_h*jf_hb)
rs("hb")=rs("hb")+hb_h
rs.update
response.write "<script language=JavaScript>" & chr(13) & "alert('转换成功！');" &"window.location='props.asp?"&C_ID&"'" & "</script>"
Call CloseRs(rs)
end Sub

sub dick_3()
hb_j=trim(request.form("hb_j"))
set rs=server.createobject("adodb.recordset")
sql = "select jf,hb from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,3
if hb_j="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('请选择转换数量！');" &"window.location='props.asp?"&C_ID&"'" & "</script>"
response.end
end if
rs("hb")=rs("hb")-int(hb_j/adjfs)
rs("jf")=rs("jf")+hb_j
rs.update
response.write "<script language=JavaScript>" & chr(13) & "alert('转换成功！');" &"window.location='props.asp?"&C_ID&"'" & "</script>"
Call CloseRs(rs)
end Sub

sub dick_4()
aliname=trim(request.form("aliname"))
bbb=CInt(trim(request.form("bbb")))
Amount2=CInt(request.form("Amount2"))
czbz=request.form("czbz")
if Amount2<=0 Or Amount2<bbb*rmb_hb Then
response.write "<script language=JavaScript>" & chr(13) & "alert('您帐户的人民币余额不足，无法兑换！');" &"window.location='props.asp?"&C_ID&"'" & "</script>"
response.end
end If
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_Financial] "
rs.open sql,conn,1,3
rs.addnew
rs("username")=username
rs("aliname")=aliname
rs("ywje")=bbb*rmb_hb
rs("ywlx")="支出"
rs("czbz")=czbz
rs("czz")=username
rs("admincl")=1
rs("addTime")=now()
rs.update
Call CloseRs(rs)
conn.execute "UPDATE DNJY_user SET hb=hb+"&bbb*rmb_hb&" WHERE username='"&username&"'" '同时向用户更新帐户虚拟币数量
conn.execute "UPDATE DNJY_user SET Amount=Amount-"&bbb*rmb_hb&" WHERE username='"&username&"'" '同时扣除相应人民币
response.write "<script language=JavaScript>" & chr(13) & "alert('操作成功！');" &"window.location='props.asp?"&C_ID&"'" & "</script>"
end Sub

sub dick_5()
aliname=trim(request.form("aliname"))
hb=CInt(trim(request.form("hb")))
Amount2=CInt(request.form("Amount2"))
czbz=request.form("czbz")
if hb<=0 Or hb<Amount2*rmb_hb Then
response.write "<script language=JavaScript>" & chr(13) & "alert('您帐户的"&rmb_mc&"余额不足，无法兑换！');" &"window.location='props.asp?"&C_ID&"'" & "</script>"
response.end
end If
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_Financial] "
rs.open sql,conn,1,3
rs.addnew
rs("username")=username
rs("aliname")=aliname
rs("ywje")=Amount2
rs("ywlx")="入帐"
rs("czbz")=czbz
rs("czz")=username
rs("admincl")=1
rs("addTime")=now()
rs.update
Call CloseRs(rs)
conn.execute "UPDATE DNJY_user SET Amount=Amount+"&Amount2&" WHERE username='"&username&"'" '同时向用户更新帐户金额
conn.execute "UPDATE DNJY_user SET hb=hb-"&Amount2*rmb_hb&" WHERE username='"&username&"'" '同时扣除相应虚拟币
response.write "<script language=JavaScript>" & chr(13) & "alert('操作成功！');" &"window.location='props.asp?"&C_ID&"'" & "</script>"
end Sub
%>

        <tr>
          <td width="94%" height="1">　<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="26">
            <tr>
              <td width="100%" height="26">　
              <%
              if request("dick")="1" then
              dick_1()
              elseif request("dick")="2" then
              dick_2()
              elseif request("dick")="3" then
              dick_3()
			  elseif request("dick")="4" then
              dick_4()
              elseif request("dick")="5" then
              dick_5()
              end if
              %>
              </td>
            </tr>
          </table>
          </td>
        </tr>

        <tr>
          <td width="99%" height="25" align="center">
          <p align="left">　</td>
        </tr>
        <tr>
          <td width="99%" height="25" align="center">
          <p align="left">　</td>
        </tr>
        <tr>
          <td width="99%" height="25">　</td>
        </tr>
        </table>
      </td>
      <td width="12" height="299" background="img/line_01.gif">　</td>
    </tr>
    <tr>
      <td height="26" colspan="4"><!--#include file=footer.asp--></td>
    </tr>
  </table>
  </center>
</div>
</body>
</html>

