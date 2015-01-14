<!--#include file="../conn.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file="../Include/Form_board .asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Dim Action,ID,Fso,Ts,i,Moldboard,Text,TextFile,m_name,temp,M_logo,mb,m_,data,M(96)'当前最大变量，实际已用到96
Moldboard=Request.Form("Moldboard")
ID=strint(request.QueryString("ID"))
Call OpenConn
IF Request.Form("Submit")<>"" Then
For i=0 to 96'当前最大变量
IF i=0 Then
m(i)=request.form("M"&i)
Else
m(i)=strint(Request.Form("M"&i))
End IF
Temp=Temp&m(i)
IF i<>96 Then Temp=Temp&"|"'当前最大变量
Next

Set Rs=server.createobject("adodb.recordset")
Sql="Select ID,M_list,M_liststr From DNJY_Template Where ID="&ID
Rs.Open Sql,conn,1,3
Rs(1)=Moldboard
Rs(2)=Temp
Rs.Update

set fso = CreateObject("Scripting.FileSystemObject")
set Ts = Fso.OpenTextFile(Server.Mappath("../list.asp"),2, True)
Ts.Write M_list(Moldboard,Temp)
Ts.Close:Set Ts=Nothing
Set Fso=Nothing
If Htmlhome=1 Then Call HomeHtml()'重新生成首页
response.write "<script language='javascript'>	alert('模板设置成功，并已生成新的页面！');location.href='javascript:history.go(-1)';</script>"
End IF

Set Rs=Conn.Execute("Select M_list,M_liststr,M_Name,M_logo From DNJY_Template Where ID="&ID)
IF Not Rs.Eof Then Data=Rs.GetRows
Rs.close:Set Rs=Nothing
Call CloseDB(conn)
M_logo=Data(3,0)
M_=Split(Data(1,0),"|")%>
<HTML>
<HEAD>
<LINK rel="stylesheet" type="text/css" href="../css/style.css">
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<META name="GENERATOR" content="Microsoft FrontPage 6.0">
<META name="ProgId" content="FrontPage.Editor.Document">
<TITLE>信息栏目模板</title>
<style type="text/css">
<!--
.STYLE1 {font-weight: bold}
-->
</style>
</HEAD>
<BODY background="images/background.gif">                                                                 
  <TABLE width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
    <TR>
      <TD height="20" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">信 息 栏 目 模 板 管 理
	  </FONT></TD>
    </TR>
    <TR>
      <TD bgcolor="#FFFFFF"> <BR>
            <TABLE width="98%" border="0"  align=center cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
  <FORM name="form1" method="post">
    <TR bgcolor="#FFFFFF">
      <TD width="20%" height="26" align="right">模板代码<BR>
       </TD>
      <TD width="60%" >当前模板：信息栏目模板：list.asp,当前模板方案：<FONT color="#ff0000"><%=Data(2,0)%></font>&nbsp;&nbsp;模板标识：<%=M_logo%><br>
        &nbsp;<TEXTAREA name="moldboard" cols="90" rows="25"><%=server.HTMLEncode(Data(0,0))%></TEXTAREA>
        <BR>
        <BR>
        </P></TD>
	<TD width="20%" height="26" align="left"> 
        模板参数：<BR>
<FONT color="#ff0000">
{$搜索条}<br>
{$信息分类列表}<br>
{$地区分类列表}<br>
{$最新信息}<br>
{$竞价信息}<br>
{$推荐信息}<br>
{$热门信息}<br>
{$栏目信息}<br>
{$信息图文列表}<br></FONT>
<br> <br><a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')"><b>0-15种日期样式</b></a> </TD>
      </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right"><b>注意：</b></TD>
      <TD colspan="2"><b>以下打[√]为当前模板有效,设置参数有效，打[<FONT color=#ff0000>×</font>]为当前模板无效,设置参数无效！</b></TD>
    </TR>
	<TR bgcolor="#FFFFFF">
      <TD height="30" align="right"  bgcolor="#BDDEEF">站内搜索条：<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2" bgcolor="#BDDEEF">&nbsp;分类显示：
        <SELECT name="m0" id="m0">
          <OPTION value="1" <%IF StrInt(m_(0))=1 Then%>selected<%End IF%>>显示一级(推荐)</OPTION>
          <OPTION value="2" <%IF StrInt(m_(0))=2 Then%>selected<%End IF%>>显示一级和二级</OPTION>
          <OPTION value="3" <%IF StrInt(m_(0))=3 Then%>selected<%End IF%>>显示全部(不推荐)</OPTION>
        </SELECT>
        地区显示：
        <SELECT name="m1" id="m1">
          <OPTION value="1" <%IF StrInt(m_(1))=1 Then%>selected<%End IF%>>显示一级(推荐)</OPTION>
          <OPTION value="2" <%IF StrInt(m_(1))=2 Then%>selected<%End IF%>>显示一级和二级</OPTION>
          <OPTION value="3" <%IF StrInt(m_(1))=3 Then%>selected<%End IF%>>显示全部(不推荐)</OPTION>
        </SELECT>（因全部地区列表速度太慢，地区显示默认显示一级，此项选择无效）</TD>
    </TR>

    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">信息分类显示设置：<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2">
	        标题圆点标志：<INPUT name="m2" type="checkbox" value="1" <%IF strint(m_(2))=1 Then%>checked<%End IF%>>
			属下信息条数：<INPUT name="m3" type="checkbox" value="1" <%IF strint(m_(3))=1 Then%>checked<%End IF%>>            
			颜色显示：<INPUT name="m4" type="checkbox" value="1" <%IF strint(m_(4))=1 Then%>checked<%End IF%>> 
            显示列数：<INPUT name="m5" type="text" value="<%=m_(5)%>" size="6">
            行距：<INPUT name="m6" type="text" value="<%=m_(6)%>" size="5">px
</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" bgcolor="#BDDEEF">地区分类显示设置：<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2" bgcolor="#BDDEEF">
	        标题圆点标志：<INPUT name="m7" type="checkbox" value="1" <%IF strint(m_(7))=1 Then%>checked<%End IF%>>
			属下信息条数：<INPUT name="m8" type="checkbox" value="1" <%IF strint(m_(8))=1 Then%>checked<%End IF%>>            
			颜色显示：<INPUT name="m9" type="checkbox" value="1" <%IF strint(m_(9))=1 Then%>checked<%End IF%>> 
            显示列数：<INPUT name="m10" type="text" value="<%=m_(10)%>" size="6">
            行距：<INPUT name="m11" type="text" value="<%=m_(11)%>" size="5">px</TD>
    </TR>

    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">最新信息显示设置：<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2">
	        信息类型：<INPUT name="m12" type="checkbox" value="1" <%IF strint(m_(12))=1 Then%>checked<%End IF%>>
			推荐标志：<INPUT name="m13" type="checkbox" value="1" <%IF strint(m_(13))=1 Then%>checked<%End IF%>>
			启用固顶：<INPUT name="m14" type="checkbox" value="1" <%IF strint(m_(14))=1 Then%>checked<%End IF%>>
           	图片标志：<INPUT name="m15" type="checkbox" value="1" <%IF strint(m_(15))=1 Then%>checked<%End IF%>>
			日期样式(0-15)：<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">【例】</a><INPUT name="m16" type="text" value="<%=m_(16)%>" size="5">			
			竞价提示：<INPUT name="m17" type="checkbox" value="1" <%IF strint(m_(17))=1 Then%>checked<%End IF%>>
			显示更多：<INPUT name="m18" type="checkbox" value="1" <%IF strint(m_(18))=1 Then%>checked<%End IF%>>
			标题颜色：<INPUT name="m19" type="checkbox" value="1" <%IF strint(m_(19))=1 Then%>checked<%End IF%>><br>
			信笺背景：<INPUT name="m20" type="checkbox" value="1" <%IF strint(m_(20))=1 Then%>checked<%End IF%>>
			点击数量：<INPUT name="m21" type="checkbox" value="1" <%IF strint(m_(21))=1 Then%>checked<%End IF%>>

标题字数：<INPUT name="m22" type="text" value="<%=m_(22)%>" size="5">
显示条数：<INPUT name="m23" type="text" value="<%=m_(23)%>" size="5"> 
显示列数：<INPUT name="m24" type="text" value="<%=m_(24)%>" size="6">
字体大小：<INPUT name="m25" type="text" value="<%=m_(25)%>" size="5">px
 行距：<INPUT name="m26" type="text" value="<%=m_(26)%>" size="5">px</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" bgcolor="#BDDEEF">竞价信息显示设置：<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2" bgcolor="#BDDEEF">
	        信息类型：<INPUT name="m27" type="checkbox" value="1" <%IF strint(m_(27))=1 Then%>checked<%End IF%>>
			推荐标志：<INPUT name="m28" type="checkbox" value="1" <%IF strint(m_(28))=1 Then%>checked<%End IF%>>
			启用固顶：<INPUT name="m29" type="checkbox" value="1" <%IF strint(m_(29))=1 Then%>checked<%End IF%>>
           	图片标志：<INPUT name="m30" type="checkbox" value="1" <%IF strint(m_(30))=1 Then%>checked<%End IF%>>
			日期样式(0-15)：<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">【例】</a><INPUT name="m31" type="text" value="<%=m_(31)%>" size="5">
			竞价提示：<INPUT name="m32" type="checkbox" value="1" <%IF strint(m_(32))=1 Then%>checked<%End IF%>>
			显示更多：<INPUT name="m33" type="checkbox" value="1" <%IF strint(m_(33))=1 Then%>checked<%End IF%>>
			标题颜色：<INPUT name="m34" type="checkbox" value="1" <%IF strint(m_(34))=1 Then%>checked<%End IF%>><br>
			信笺背景：<INPUT name="m35" type="checkbox" value="1" <%IF strint(m_(35))=1 Then%>checked<%End IF%>>
			点击数量：<INPUT name="m36" type="checkbox" value="1" <%IF strint(m_(36))=1 Then%>checked<%End IF%>>
			
标题字数：<INPUT name="m37" type="text" value="<%=m_(37)%>" size="5">
显示条数：<INPUT name="m38" type="text" value="<%=m_(38)%>" size="5"> 
显示列数：<INPUT name="m39" type="text" value="<%=m_(39)%>" size="6">
字体大小：<INPUT name="m40" type="text" value="<%=m_(40)%>" size="5">px
 行距：<INPUT name="m41" type="text" value="<%=m_(41)%>" size="5">px
	        </TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">推荐信息显示设置：<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2">
	        信息类型：<INPUT name="m42" type="checkbox" value="1" <%IF strint(m_(42))=1 Then%>checked<%End IF%>>
			推荐标志：<INPUT name="m43" type="checkbox" value="1" <%IF strint(m_(43))=1 Then%>checked<%End IF%>>
			启用固顶：<INPUT name="m44" type="checkbox" value="1" <%IF strint(m_(44))=1 Then%>checked<%End IF%>>
           	图片标志：<INPUT name="m45" type="checkbox" value="1" <%IF strint(m_(45))=1 Then%>checked<%End IF%>>
			日期样式(0-15)：<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">【例】</a><INPUT name="m46" type="text" value="<%=m_(46)%>" size="5">			
			竞价提示：<INPUT name="m47" type="checkbox" value="1" <%IF strint(m_(47))=1 Then%>checked<%End IF%>>
			显示更多：<INPUT name="m48" type="checkbox" value="1" <%IF strint(m_(48))=1 Then%>checked<%End IF%>>
			标题颜色：<INPUT name="m49" type="checkbox" value="1" <%IF strint(m_(49))=1 Then%>checked<%End IF%>><br>
			信笺背景：<INPUT name="m50" type="checkbox" value="1" <%IF strint(m_(50))=1 Then%>checked<%End IF%>>
			点击数量：<INPUT name="m51" type="checkbox" value="1" <%IF strint(m_(51))=1 Then%>checked<%End IF%>>

标题字数：<INPUT name="m52" type="text" value="<%=m_(52)%>" size="5">
显示条数：<INPUT name="m53" type="text" value="<%=m_(53)%>" size="5"> 
显示列数：<INPUT name="m54" type="text" value="<%=m_(54)%>" size="6">
字体大小：<INPUT name="m55" type="text" value="<%=m_(55)%>" size="5">px
 行距：<INPUT name="m56" type="text" value="<%=m_(56)%>" size="5">px
</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" bgcolor="#BDDEEF">热门信息显示设置：<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2" bgcolor="#BDDEEF">
	        信息类型：<INPUT name="m57" type="checkbox" value="1" <%IF strint(m_(57))=1 Then%>checked<%End IF%>>
			推荐标志：<INPUT name="m58" type="checkbox" value="1" <%IF strint(m_(58))=1 Then%>checked<%End IF%>>
			启用固顶：<INPUT name="m59" type="checkbox" value="1" <%IF strint(m_(59))=1 Then%>checked<%End IF%>>
           	图片标志：<INPUT name="m60" type="checkbox" value="1" <%IF strint(m_(60))=1 Then%>checked<%End IF%>>
			日期样式(0-15)：<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">【例】</a><INPUT name="m61" type="text" value="<%=m_(61)%>" size="5">			
			竞价提示：<INPUT name="m62" type="checkbox" value="1" <%IF strint(m_(62))=1 Then%>checked<%End IF%>>
			显示更多：<INPUT name="m63" type="checkbox" value="1" <%IF strint(m_(63))=1 Then%>checked<%End IF%>>
			标题颜色：<INPUT name="m64" type="checkbox" value="1" <%IF strint(m_(64))=1 Then%>checked<%End IF%>><br>
			信笺背景：<INPUT name="m65" type="checkbox" value="1" <%IF strint(m_(65))=1 Then%>checked<%End IF%>>
			点击数量：<INPUT name="m66" type="checkbox" value="1" <%IF strint(m_(66))=1 Then%>checked<%End IF%>>

标题字数：<INPUT name="m67" type="text" value="<%=m_(67)%>" size="5">
显示条数：<INPUT name="m68" type="text" value="<%=m_(68)%>" size="5"> 
显示列数：<INPUT name="m69" type="text" value="<%=m_(69)%>" size="6">
字体大小：<INPUT name="m70" type="text" value="<%=m_(70)%>" size="5">px
 行距：<INPUT name="m71" type="text" value="<%=m_(71)%>" size="5">px
</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">分类信息列表显示设置：<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2">
	        信息类型：<INPUT name="m72" type="checkbox" value="1" <%IF strint(m_(72))=1 Then%>checked<%End IF%>>
			图文显示：<INPUT name="m73" type="checkbox" value="1" <%IF strint(m_(73))=1 Then%>checked<%End IF%>>            
每页条数：<INPUT name="m74" type="text" value="<%=m_(74)%>" size="5">
图片宽度：<INPUT name="m75" type="text" value="<%=m_(75)%>" size="5">
图片高度：<INPUT name="m76" type="text" value="<%=m_(76)%>" size="5">
标题长度：<INPUT name="m77" type="text" value="<%=m_(77)%>" size="5">
正文截取长度：<INPUT name="m78" type="text" value="<%=m_(78)%>" size="5"> 
</TD>
    </TR>

    <TR bgcolor="#BDDEEF">
      <TD height="26" align="right">信息图文显示设置(M)：<%mb="F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2">
	        信息类型：<INPUT name="m79" type="checkbox" value="1" <%IF strint(m_(79))=1 Then%>checked<%End IF%>>
			启用固顶：<INPUT name="m80" type="checkbox" value="1" <%IF strint(m_(80))=1 Then%>checked<%End IF%>>
			推荐标志：<INPUT name="m81" type="checkbox"  value="1" <%IF strint(m_(81))=1 Then%>checked<%End IF%>>
           	图片标志：<INPUT name="m82" type="checkbox"  value="1" <%IF strint(m_(82))=1 Then%>checked<%End IF%>>						
			点击数量：<INPUT name="m83" type="checkbox"  value="1" <%IF strint(m_(83))=1 Then%>checked<%End IF%>>
			显示日期(0-15)：<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">【例】</a><INPUT name="m84" type="text" value="<%=m_(84)%>" size="5">
市场价格：<INPUT name="m85" type="checkbox"  value="1" <%IF strint(m_(85))=1 Then%>checked<%End IF%>>
本站价格：<INPUT name="m86" type="checkbox"  value="1" <%IF strint(m_(86))=1 Then%>checked<%End IF%>><br>
标题长度：<INPUT name="m87" type="text" value="<%=m_(87)%>" size="5">
标题颜色：<INPUT name="m88" type="checkbox" value="1" <%IF strint(m_(88))=1 Then%>checked<%End IF%>>
字体大小：<INPUT name="m89" type="text" value="<%=m_(89)%>" size="5">px
图片宽度：<INPUT name="m90" type="text" value="<%=m_(90)%>" size="6">
图片高度：<INPUT name="m91" type="text" value="<%=m_(91)%>" size="6">
单元宽度：<INPUT name="m92" type="text" value="<%=m_(92)%>" size="6">
单元高度：<INPUT name="m93" type="text" value="<%=m_(93)%>" size="6"><br>
每页个数：<INPUT name="m94" type="text" value="<%=m_(94)%>" size="5"> 
显示列数：<INPUT name="m95" type="text" value="<%=m_(95)%>" size="6">
显示翻页：<INPUT name="m96" type="checkbox" value="1" <%IF strint(m_(96))=1 Then%>checked<%End IF%>>

    </TR>
    <TR bgcolor="#FFFFFF"> 
      <TD height="30" align=left></TD>
      <TD height="30" colspan="2" align=left>
	  <INPUT type="submit" name="Submit" value="  设 置  ">
        <input name="Submit1" type="button" id="Submit1" value="返回" onClick="location='admin_Template_Management.asp?id=<%=ID%>'"/>&nbsp;&nbsp;&nbsp;&nbsp;<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')"><b>0-15种日期样式</b></a> </TD>
    </TR>
  </FORM>
</TABLE>
      <BR>      </TD>
    </TR>
  </TABLE>
</BODY>                    
                    
</HTML>                    
