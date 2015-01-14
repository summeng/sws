<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=usercookies.asp-->
<!--#include file=top.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
dim i_1,i_2,i_3,i_4,hb,a,b,ct,d,a_1,b_1,c_1,d_1
if request("hb")<>"" then
hb=trim(request("hb"))
end if
a=trim(request("a"))
b=trim(request("b"))
ct=trim(request("ct"))
d=trim(request("d"))
if a="" then
a=0
end if
if b="" then
b=0
end if
if ct="" then
ct=0
end if
if d="" then
d=0
end if
if request.cookies("a_1")="" then
Response.cookies("a_1")=0
else
Response.cookies("a_1")=int(request.cookies("a_1"))+int(a)
end if
if request.cookies("b_1")="" then
Response.cookies("b_1")=0
else
Response.cookies("b_1")=int(request.cookies("b_1"))+int(b)
end if
if request.cookies("c_1")="" then
Response.cookies("c_1")=0
else
Response.cookies("c_1")=int(request.cookies("c_1"))+int(ct)
end if
if request.cookies("d_1")="" then
Response.cookies("d_1")=0
else
Response.cookies("d_1")=int(request.cookies("d_1"))+int(d)
end if
username=request.cookies("dick")("username")
Call OpenConn
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select hb from [DNJY_user] where username='"&username&"'" 
rs.open sql,conn,1,1
if request("dick")="1" then
hb=hb-(a*g_a+b*g_b+ct*g_c+d*g_d)
else
Response.cookies("a_1")=0
Response.cookies("b_1")=0
Response.cookies("c_1")=0
Response.cookies("d_1")=0
hb=rs("hb")
end if
%>
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
<style type="text/css">
<!--
.style5 {color: #FF0000; font-weight: bold; }
-->
</style>
</head>

<body topmargin="0" leftmargin="0">

<div align="center">
  <center>
  <table width="760" height="371" border="0" cellpadding="0" cellspacing="0" bordercolor="#111111" bgcolor="#FFFFFF" style="border-collapse: collapse">
    <tr>
      <td width="214" height="299" valign="top"><div align="center">
            <!--#include file=userleft.asp-->　
        </div></td>
		<td width="5">&nbsp;</td>
      <td width="760" height="299" align="center" valign="top"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" height="397" class="ty1">
        <tr>
          <td height="1" colspan="6"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="25">
              <tr>
                <td width="100%" height="25" bgcolor="#F3F3F3"><span class="style5">　购买道具：</span></td>
              </tr>
              <tr>
                <td height="1" bgcolor="#CCCCCC"></td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td height="25" align="center" colspan="6"><p align="left">　当前我的<%=rmb_mc%>：<%=rs("hb")%> 个，购买道具需要你有足够的<%=rmb_mc%>才可以，如果没有请先 <a href="user_zhcz.asp?<%=C_ID%>"><font color="#0000FF">充值</font></a> 再 <a href="props.asp?<%=C_ID%>"><font color="#0000FF">购买<%=rmb_mc%></font></a></td>
        </tr>
        <tr>
          <td height="10" colspan="6" align="center"><p align="left"></td>
        </tr>
        <tr>
          <td height="25" align="left" width="12"><p align="right"><img src="images/form2_r3_c3.gif"></td>
          <td width="179" height="25" align="left"> 购买<font color="#0000FF">标题变色</font>道具</td>
          <td height="25" align="left" colspan="2"><select name="a" onChange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}">
              <option value="0">请选择购买的数量</option>
              <%
			i_1=hb/g_a
			for i=1 to int(i_1)
			%>
              <option value="?a=<%=i%>&dick=1&hb=<%=hb%>&a_1=<%=a_1%>&<%=C_ID%>"><%=i%>个</option>
              <%next%>
            </select>
          </td>
          <td width="130" height="25" align="right"><font color="#0000FF"><%=g_a%></font> 个<%=rmb_mc%>/1个道具</td>
          <td width="47" height="25" align="left">　</td>
        </tr>
        <tr>
          <td height="25" align="left" width="12"><p align="right"><img src="images/form2_r3_c3.gif"></td>
          <td width="179" height="25" align="left"> 购买<font color="#0000FF">信息置顶</font>道具</td>
          <td height="25" align="left" colspan="2"><select name="b" onChange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}">
              <option value="0">请选择购买的数量</option>
              <%
			i_2=hb/g_b
			for i=1 to int(i_2)
			%>
              <option value="?b=<%=i%>&dick=1&hb=<%=hb%>&b_1=<%=b_1%>&<%=C_ID%>"><%=i%>个</option>
              <%next%>
            </select>
          </td>
          <td width="130" height="25" align="right"><font color="#0000FF"><%=g_b%></font> 个<%=rmb_mc%>/1个道具</td>
          <td width="47" height="25" align="left"> 　</td>
        </tr>
        <tr>
          <td height="25" align="left" width="12"><p align="right"><img src="images/form2_r3_c3.gif"></td>
          <td width="179" height="25" align="left"> 购买<font color="#0000FF">内容贴图</font>道具</td>
          <td height="25" align="left" colspan="2"><select name="ct" onChange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}">
              <option value="0">请选择购买的数量</option>
              <%
			i_3=hb/g_c
			for i=1 to int(i_3)
			%>
              <option value="?ct=<%=i%>&dick=1&hb=<%=hb%>&c_1=<%=c_1%>&<%=C_ID%>"><%=i%>个</option>
              <%next%>
          </select></td>
          <td width="130" height="25" align="right"><font color="#0000FF"><%=g_c%></font> 个<%=rmb_mc%>/1个道具</td>
          <td width="47" height="25" align="left"> 　 </td>
        </tr>
        <tr>
          <td height="25" align="left" width="12"><p align="right"><img src="images/form2_r3_c3.gif"></td>
          <td width="179" height="25" align="left"> 购买<font color="#0000FF">通过验证</font>道具</td>
          <td height="25" align="left" colspan="2"><select name="d" onChange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}">
              <option value="0">请选择购买的数量</option>
              <%
			i_4=hb/g_d
			for i=1 to int(i_4)
			%>
              <option value="?d=<%=i%>&dick=1&hb=<%=hb%>&d_1=<%=d_1%>&<%=C_ID%>"><%=i%>个</option>
              <%next%>
          </select></td>
          <td width="130" height="25" align="right"><font color="#0000FF"><%=g_d%></font> 个<%=rmb_mc%>/1个道具</td>
          <td width="47" height="25" align="left"> 　</td>
        </tr>
        <tr>
          <td height="1" colspan="6"></td>
        </tr>
        <tr>
          <td width="12" height="10">　</td>
          <td width="179" height="10">　</td>
          <td height="10" colspan="2">　</td>
          <td height="10" colspan="2">　</td>
        </tr>
        <tr bgcolor="#EFEFEF">
          <td height="25" colspan="2" bordercolor="#CCCCCC" style="border-bottom-style: solid; border-bottom-width: 1"><p> <font color="#FF0000"><span class="style5">　当前购买的道具： －－－&gt;</span></font></td>
          <td height="25" colspan="2" bordercolor="#CCCCCC" style="border-bottom-style: solid; border-bottom-width: 1"><p align="left"><a href="user_props.asp?<%=C_ID%>"><font color="#0000FF">重新购买</font></a></td>
          <td height="25" colspan="2" bordercolor="#CCCCCC" style="border-bottom-style: solid; border-bottom-width: 1">　</td>
        </tr>
        <form action="user_props_save.asp?<%=C_ID%>" method="POST">
        <tr>
          <td width="54" height="25" style="border-top-style: solid; border-top-width: 1">　</td>
          <td width="115" height="25" style="border-top-style: solid; border-top-width: 1">标题道具</td>
          <td width="54" height="25" align="right" ><%=request.cookies("a_1")%><input type="hidden" name="a" size="6" value="<%=request.cookies("a_1")%>">　</td>
          <td width="112" height="25" >&nbsp;个</td>
          <td width="240" height="25" colspan="2" style="border-top-style: solid; border-top-width: 1">　</td>
        </tr>
        <tr>
          <td width="54" height="25">　</td>
          <td width="115" height="25">信息置顶</td>
          <td width="54" height="25" align="right"><%=request.cookies("b_1")%><input type="hidden" name="b" size="6" value="<%=request.cookies("b_1")%>">　</td>
          <td width="112" height="25">&nbsp;个</td>
          <td width="240" height="25" colspan="2">　</td>
        </tr>
        <tr>
          <td width="54" height="25">　</td>
          <td width="115" height="25">内容贴图</td>
          <td width="54" height="25" align="right"><%=request.cookies("c_1")%><input type="hidden" name="ct" size="6" value="<%=request.cookies("c_1")%>">　</td>
          <td width="112" height="25">&nbsp;个</td>
          <td width="240" height="25" colspan="2">
          <input type="submit" value="确定购买" name="B1"></td>
        </tr>
        <tr>
          <td width="54" height="25" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#CCCCCC">　</td>
          <td width="115" height="25" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#CCCCCC">通过验证</td>
          <td width="54" height="25" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#CCCCCC" align="right"><%=request.cookies("d_1")%><input type="hidden" name="d" size="6" value="<%=request.cookies("d_1")%>">　</td>
          <td width="112" height="25" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#CCCCCC">&nbsp;个</td>
          <td width="240" height="25" colspan="2" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#CCCCCC">　</td>
        </tr>
        <tr>
          <td width="575" height="25" style="border-top-style: solid; border-top-width: 1" colspan="6">
          <font color="#FF0000">注意：在该页购买道具时不要重复刷新页面，否则会导致购买道具的数量混乱，且支付会不成功。</font></td>
        </tr>
        </form>
      </table></td>
      <td width="4" align="center" valign="top" bgcolor="#e6e6e6"></td>
    </tr>
    <tr><%Call CloseRs(rs)%>
      <td height="26" colspan="3"><!--#include file=footer.asp--></td>
    </tr>
  </table>
  </center>
</div>
</body>
</html>

