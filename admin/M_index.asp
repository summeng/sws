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
<%
Dim Action,ID,fso,Ts,Data,Moldboard,i,Temp,TextFile,Text,M_logo,mb,M_,M(295)'当前实际用到的最大变量
Moldboard=Request.Form("Moldboard")
ID=strint(request.QueryString("ID"))
Call OpenConn
IF Request.Form("Submit")<>"" Then
For i=0 to 295'当前最大变量
IF i=176 Then'企业文字名片标题颜色
m(i)=HtmlEncode2(Request.Form("M"&i))
Else
m(i)=strint(Request.Form("M"&i))
End IF
Temp=Temp&m(i)
IF i<>295 Then Temp=Temp&"|"'当前最大变量
Next

Set Rs=server.createobject("adodb.recordset")
Sql="Select ID,M_Index,M_Indexstr From DNJY_Template Where ID="&ID
Rs.Open Sql,conn,1,3
Rs(1)=Moldboard
Rs(2)=Temp
Rs.Update

set fso = CreateObject("Scripting.FileSystemObject")
set Ts = Fso.OpenTextFile(Server.Mappath("../Index.asp"),2, True)
Ts.Write M_Index(Moldboard,Temp)
Ts.Close:Set Ts=Nothing
Set Fso=Nothing
If Htmlhome=1 Then Call HomeHtml()'重新生成首页
response.write "<script language='javascript'>	alert('模板设置成功，并已生成新的页面\n\n请重新生成静态内容页！');location.href='javascript:history.go(-1)';</script>"
End IF

Set Rs=Conn.Execute("Select M_Index,M_Indexstr,M_Name,M_logo From DNJY_Template Where ID="&ID)
IF Not Rs.Eof Then Data=Rs.GetRows
Rs.close:Set Rs=Nothing
Call CloseDB(conn)
M_logo=Data(3,0)
M_=Split(Data(1,0),"|")
%>
<HTML>
<HEAD>
<LINK rel="stylesheet" type="text/css" href="../css/style.css">
<SCRIPT language="javascript" src="admin.js"></SCRIPT>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<META name="GENERATOR" content="Microsoft FrontPage 6.0">
<META name="ProgId" content="FrontPage.Editor.Document">
<TITLE>首页模板</title>
<style type="text/css">
<!--
body {
	background-color: #D6DFF7;
}
body,td,th {
	color: #135294;
}
-->
</style></HEAD>
<BODY background="images/background.gif">  
<IFRAME width="260" height="165" id="colourPalette" src="nc_selcolor.htm" style="visibility:hidden; position: absolute; left: 0px; top: 0px;border:1px gray solid" frameborder="0" scrolling="no" ></IFRAME>                                                               
  <TABLE width="98%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#799AE1">
    <TR>
      <TD height="20" align="center" bgcolor="#799AE1"><FONT color="#FFFFFF" style="font-size:14px">首 页 模 板 管 理</FONT></TD>
    </TR>
    <TR>
      <TD bgcolor="#FFFFFF">
            <TABLE width="98%" border="0"  align=center cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
  <FORM name="form1" method="post" action="">
    <TR bgcolor="#FFFFFF">
      <TD width="20%" height="26" align="right">模板代码</TD>
      <TD width="60%" bgcolor="#FFFFFF" valign="top">当前模板：首页模板index.asp，当前模板方案：<FONT color="#ff0000"><%=Data(2,0)%></font>
        &nbsp;&nbsp;模板标识：<%=M_logo%><br><TEXTAREA name="Moldboard" cols="90" rows="30" id="T2"><%=Data(0,0)%></TEXTAREA></TD>
		<TD width="20%" height="26" align="left" >模版标签：<br>
      可通过如下标签调用相关内容：<p>
<FONT color="#ff0000">{$搜索条}<br>
{$最新信息}<br>
{$竞价信息}<br>
{$推荐信息}<br>
{$热门信息}<br>
{$分类信息}<br>
{$方框显示信息}<br>
{$图片信息推荐}<br>
{$最新信息图文}<br>
{$推荐信息图文}<br>
{$热门信息图文}<br>
{$名店推荐}<br>
{$店铺资源共享}<br>
{$企业文字名片}<br>
{$企业图片名片}<br>
{$新闻资讯一}<br>
{$新闻资讯二}<br>
{$新闻资讯三}<br>
{$新闻资讯四}<br>
{$新闻资讯相册}<br>
{$电子杂志}<br>
{$都市114}<br>
{$便民服务}<br>
{$友情链接}<br>
{$用户留言}<br>
{$本站公告}<br>
</font><br> <br>
<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')"><b>0-15种显示日期</b></a> 
</TD></TR>
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
      <TD colspan="2" bgcolor="#BDDEEF">&nbsp;信息分类显示：
        <SELECT name="m0">
          <OPTION value="1" <%IF StrInt(m_(0))=1 Then%>selected<%End IF%>>显示一级(推荐)</OPTION>
          <OPTION value="2" <%IF StrInt(m_(0))=2 Then%>selected<%End IF%>>显示一级和二级</OPTION>
          <OPTION value="3" <%IF StrInt(m_(0))=3 Then%>selected<%End IF%>>显示全部(不推荐)</OPTION>
        </SELECT>
        地区分类显示：
        <SELECT name="m1">
          <OPTION value="1" <%IF StrInt(m_(1))=1 Then%>selected<%End IF%>>显示一级(推荐)</OPTION>
          <OPTION value="2" <%IF StrInt(m_(1))=2 Then%>selected<%End IF%>>显示一级和二级</OPTION>
          <OPTION value="3" <%IF StrInt(m_(1))=3 Then%>selected<%End IF%>>显示全部(不推荐)</OPTION>
        </SELECT>（若地区分类太多列表速度会很慢，看情况选择显示多少级）</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">最新信息显示设置：<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2">
	        信息类型：<INPUT name="m2" type="checkbox" value="1" <%IF strint(m_(2))=1 Then%>checked<%End IF%>>
			启用固顶：<INPUT name="m3" type="checkbox" value="1" <%IF strint(m_(3))=1 Then%>checked<%End IF%>>
			推荐标志：<INPUT name="m4" type="checkbox" value="1" <%IF strint(m_(4))=1 Then%>checked<%End IF%>>
           	图片标志：<INPUT name="m5" type="checkbox" value="1" <%IF strint(m_(5))=1 Then%>checked<%End IF%>>			
			竞价提示：<INPUT name="m6" type="checkbox" value="1" <%IF strint(m_(6))=1 Then%>checked<%End IF%>>
			显示更多：<INPUT name="m7" type="checkbox" value="1" <%IF strint(m_(7))=1 Then%>checked<%End IF%>>
			标题颜色：<INPUT name="m8" type="checkbox" value="1" <%IF strint(m_(8))=1 Then%>checked<%End IF%>>
			信笺背景：<INPUT name="m9" type="checkbox" value="1" <%IF strint(m_(9))=1 Then%>checked<%End IF%>>
			点击数量：<INPUT name="m10" type="checkbox" value="1" <%IF strint(m_(10))=1 Then%>checked<%End IF%>><br>
			显示日期(0-15)：<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">【例】</a>
			<INPUT name="m11" type="text" value="<%=m_(11)%>" size="5">
标题字数：<INPUT name="m12" type="text" value="<%=m_(12)%>" size="5">
显示条数：<INPUT name="m13" type="text" value="<%=m_(13)%>" size="5"> 
显示列数：<INPUT name="m14" type="text" value="<%=m_(14)%>" size="6">
字体大小：<INPUT name="m15" type="text" value="<%=m_(15)%>" size="5">px
 行距：<INPUT name="m16" type="text" value="<%=m_(16)%>" size="5">px</TD>
    </TR>

    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" bgcolor="#BDDEEF">竞价信息显示设置：<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2" bgcolor="#BDDEEF">
	        信息类型：<INPUT name="m17" type="checkbox" value="1" <%IF strint(m_(17))=1 Then%>checked<%End IF%>>
			启用固顶：<INPUT name="m18" type="checkbox" value="1" <%IF strint(m_(18))=1 Then%>checked<%End IF%>>
			推荐标志：<INPUT name="m19" type="checkbox" value="1" <%IF strint(m_(19))=1 Then%>checked<%End IF%>> 
           	图片标志：<INPUT name="m20" type="checkbox" value="1" <%IF strint(m_(20))=1 Then%>checked<%End IF%>>			
			竞价提示：<INPUT name="m21" type="checkbox" value="1" <%IF strint(m_(21))=1 Then%>checked<%End IF%>>
			显示更多：<INPUT name="m22" type="checkbox" value="1" <%IF strint(m_(22))=1 Then%>checked<%End IF%>>
			标题颜色：<INPUT name="m23" type="checkbox" value="1" <%IF strint(m_(23))=1 Then%>checked<%End IF%>>
			信笺背景：<INPUT name="m24" type="checkbox" value="1" <%IF strint(m_(24))=1 Then%>checked<%End IF%>>
			点击数量：<INPUT name="m25" type="checkbox" value="1" <%IF strint(m_(25))=1 Then%>checked<%End IF%>><br>
			显示日期(0-15)：<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">【例】</a>
			<INPUT name="m26" type="text" value="<%=m_(26)%>" size="5">
标题字数：<INPUT name="m27" type="text" value="<%=m_(27)%>" size="5">
显示条数：<INPUT name="m28" type="text" value="<%=m_(28)%>" size="5"> 
显示列数：<INPUT name="m29" type="text" value="<%=m_(29)%>" size="6">
字体大小：<INPUT name="m30" type="text" value="<%=m_(30)%>" size="5">px
 行距：<INPUT name="m31" type="text" value="<%=m_(31)%>" size="5">px
	        </TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">推荐信息显示设置：<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2">
	        信息类型：<INPUT name="m32" type="checkbox" value="1" <%IF strint(m_(32))=1 Then%>checked<%End IF%>>
			启用固顶：<INPUT name="m33" type="checkbox" value="1" <%IF strint(m_(33))=1 Then%>checked<%End IF%>>
			推荐标志：<INPUT name="m34" type="checkbox" value="1" <%IF strint(m_(34))=1 Then%>checked<%End IF%>>
           	图片标志：<INPUT name="m35" type="checkbox" value="1" <%IF strint(m_(35))=1 Then%>checked<%End IF%>>			
			竞价提示：<INPUT name="m36" type="checkbox" value="1" <%IF strint(m_(36))=1 Then%>checked<%End IF%>>
			显示更多：<INPUT name="m37" type="checkbox" value="1" <%IF strint(m_(37))=1 Then%>checked<%End IF%>>
			标题颜色：<INPUT name="m38" type="checkbox" value="1" <%IF strint(m_(38))=1 Then%>checked<%End IF%>>
			信笺背景：<INPUT name="m39" type="checkbox" value="1" <%IF strint(m_(39))=1 Then%>checked<%End IF%>>
			点击数量：<INPUT name="m40" type="checkbox" value="1" <%IF strint(m_(40))=1 Then%>checked<%End IF%>><br>
			显示日期(0-15)：<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">【例】</a>
			<INPUT name="m41" type="text" value="<%=m_(41)%>" size="5">
标题字数：<INPUT name="m42" type="text" value="<%=m_(42)%>" size="5">
显示条数：<INPUT name="m43" type="text" value="<%=m_(43)%>" size="5"> 
显示列数：<INPUT name="m44" type="text" value="<%=m_(44)%>" size="6">
字体大小：<INPUT name="m45" type="text" value="<%=m_(45)%>" size="5">px
 行距：<INPUT name="m46" type="text" value="<%=m_(46)%>" size="5">px
</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" bgcolor="#BDDEEF">热门信息显示设置：<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2" bgcolor="#BDDEEF">
	  	    信息类型：<INPUT name="m47" type="checkbox" value="1" <%IF strint(m_(47))=1 Then%>checked<%End IF%>>
			启用固顶：<INPUT name="m48" type="checkbox" value="1" <%IF strint(m_(48))=1 Then%>checked<%End IF%>>
			推荐标志：<INPUT name="m49" type="checkbox" value="1" <%IF strint(m_(49))=1 Then%>checked<%End IF%>> 
           	图片标志：<INPUT name="m50" type="checkbox" value="1" <%IF strint(m_(50))=1 Then%>checked<%End IF%>>			
			竞价提示：<INPUT name="m51" type="checkbox" value="1" <%IF strint(m_(51))=1 Then%>checked<%End IF%>>
			显示更多：<INPUT name="m52" type="checkbox" value="1" <%IF strint(m_(52))=1 Then%>checked<%End IF%>>
			标题颜色：<INPUT name="m53" type="checkbox" value="1" <%IF strint(m_(53))=1 Then%>checked<%End IF%>>
			信笺背景：<INPUT name="m54" type="checkbox" value="1" <%IF strint(m_(54))=1 Then%>checked<%End IF%>>
			点击数量：<INPUT name="m55" type="checkbox" value="1" <%IF strint(m_(55))=1 Then%>checked<%End IF%>><br>
			显示日期(0-15)：<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">【例】</a>
			<INPUT name="m56" type="text" value="<%=m_(56)%>" size="5">
标题字数：<INPUT name="m57" type="text" value="<%=m_(57)%>" size="5">
显示条数：<INPUT name="m58" type="text" value="<%=m_(58)%>" size="5"> 
显示列数：<INPUT name="m59" type="text" value="<%=m_(59)%>" size="6">
字体大小：<INPUT name="m60" type="text" value="<%=m_(60)%>" size="5">px
 行距：<INPUT name="m61" type="text" value="<%=m_(61)%>" size="5">px</TD>
    </TR>
    <TR bgcolor="#FFFFFF"> 
      <TD height="26" align=right>信息按分类列表显示设置：<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD width="73%" colspan="2" align=left>
	  	    信息类型：<INPUT name="m62" type="checkbox" value="1" <%IF strint(m_(62))=1 Then%>checked<%End IF%>>
			分类个数：<INPUT name="m63" type="text" value="<%=m_(63)%>" size="5">
            显示条数：<INPUT name="m64" type="text" value="<%=m_(64)%>" size="5">
            显示列数：<INPUT name="m65" type="text" value="<%=m_(65)%>" size="5">
			标题长度：<INPUT name="m66" type="text" value="<%=m_(66)%>" size="5">
			标题颜色：<INPUT name="m67" type="checkbox" value="1" <%IF strint(m_(67))=1 Then%>checked<%End IF%>>
			启用固顶：<INPUT name="m68" type="checkbox" value="1" <%IF strint(m_(68))=1 Then%>checked<%End IF%>>
	  	    图片标志：<INPUT name="m69" type="checkbox" value="1" <%IF strint(m_(69))=1 Then%>checked<%End IF%>>	  	    
	  	    点击数：<INPUT name="m70" type="checkbox" value="1" <%IF strint(m_(70))=1 Then%>checked<%End IF%>><br>			
	  	    显示日期(0-15)：<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">【例】</a>
			<INPUT name="m71" type="text" value="<%=m_(71)%>" size="5">
			头条图文：<INPUT name="m72" type="checkbox"  value="1" <%IF strint(m_(72))=1 Then%>checked<%End IF%>>	  
            信笺背景：
            <SELECT name="m73">
              <OPTION value="0" <%IF StrInt(m_(73))=0 Then%>selected<%End IF%>>要信笺背景</OPTION>
              <OPTION value="1" <%IF StrInt(m_(73))=1 Then%>selected<%End IF%>>无信笺保格式</OPTION>
			  <OPTION value="2" <%IF StrInt(m_(73))=2 Then%>selected<%End IF%>>无信笺无格式</OPTION>
            </SELECT><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT color="#ff0000">列宽要在本模板方案－CSS风格中找到shadow1，修改该段width:的值，与信息频道页共用</font></TD>
    </TR>

    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" bgcolor="#BDDEEF">竞价信息方框显示设置：<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2" bgcolor="#BDDEEF">
	        标题颜色：<INPUT name="m74" type="checkbox" value="1" <%IF strint(m_(74))=1 Then%>checked<%End IF%>>
			信息类型：<INPUT name="m75" type="checkbox" value="1" <%IF strint(m_(75))=1 Then%>checked<%End IF%>>            
			信息LOGO：<INPUT name="m76" type="checkbox" value="1" <%IF strint(m_(76))=1 Then%>checked<%End IF%>>
           	启用固顶：<INPUT name="m77" type="checkbox" value="1" <%IF strint(m_(77))=1 Then%>checked<%End IF%>>
			显示有效期：<INPUT name="m78" type="checkbox" value="1" <%IF strint(m_(78))=1 Then%>checked<%End IF%>>
			图片标志：<INPUT name="m79" type="checkbox" value="1" <%IF strint(m_(79))=1 Then%>checked<%End IF%>>
			点击数量：<INPUT name="m80" type="checkbox" value="1" <%IF strint(m_(80))=1 Then%>checked<%End IF%>>
			显示会员：<INPUT name="m81" type="checkbox" value="1" <%IF strint(m_(81))=1 Then%>checked<%End IF%>>
			显示竞价：<INPUT name="m82" type="checkbox" value="1" <%IF strint(m_(82))=1 Then%>checked<%End IF%>><br>

标题长度：<INPUT name="m83" type="text" value="<%=m_(83)%>" size="5">
简介长度：<INPUT name="m84" type="text" value="<%=m_(84)%>" size="5"> 
显示列数：<INPUT name="m85" type="text" value="<%=m_(85)%>" size="6">
显示总数：<INPUT name="m86" type="text" value="<%=m_(86)%>" size="5">
			信息范围：
            <SELECT name="m87">
              <OPTION value="0" <%IF strint(m_(87))=0 Then%>selected<%End IF%>>显示全部信息</OPTION>
              <OPTION value="1" <%IF strint(m_(87))=1 Then%>selected<%End IF%>>仅显示竞价信息</OPTION>			  
            </SELECT>
</TD>
    </TR>

    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">图片信息推荐显示：<%mb=""
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2">
每行显示条数：<INPUT name="m88" type="text" value="<%=m_(88)%>" size="5">
总共显示条数：<INPUT name="m89" type="text" value="<%=m_(89)%>" size="5">
		</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" bgcolor="#BDDEEF">最新信息图文显示设置：<%mb="F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2" bgcolor="#BDDEEF">
	        信息类型：<INPUT name="m90" type="checkbox" value="1" <%IF strint(m_(90))=1 Then%>checked<%End IF%>>
			启用固顶：<INPUT name="m91" type="checkbox" value="1" <%IF strint(m_(91))=1 Then%>checked<%End IF%>>
			推荐标志：<INPUT name="m92" type="checkbox"  value="1" <%IF strint(m_(92))=1 Then%>checked<%End IF%>>
           	图片标志：<INPUT name="m93" type="checkbox"  value="1" <%IF strint(m_(93))=1 Then%>checked<%End IF%>>						
			点击数量：<INPUT name="m94" type="checkbox"  value="1" <%IF strint(m_(94))=1 Then%>checked<%End IF%>>
			显示日期(0-15)：<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">【例】</a>
			<INPUT name="m95" type="text" value="<%=m_(95)%>" size="5">
市场价格：<INPUT name="m96" type="checkbox"  value="1" <%IF strint(m_(96))=1 Then%>checked<%End IF%>>
本站价格：<INPUT name="m97" type="checkbox"  value="1" <%IF strint(m_(97))=1 Then%>checked<%End IF%>><br>
标题长度：<INPUT name="m98" type="text" value="<%=m_(98)%>" size="5">
标题颜色：<INPUT name="m99" type="checkbox" value="1" <%IF strint(m_(99))=1 Then%>checked<%End IF%>>
字体大小：<INPUT name="m100" type="text" value="<%=m_(100)%>" size="5">px
图片宽度：<INPUT name="m101" type="text" value="<%=m_(101)%>" size="6">
图片高度：<INPUT name="m102" type="text" value="<%=m_(102)%>" size="6">
单元宽度：<INPUT name="m103" type="text" value="<%=m_(103)%>" size="6">
单元高度：<INPUT name="m104" type="text" value="<%=m_(104)%>" size="6"><br>
每页个数：<INPUT name="m105" type="text" value="<%=m_(105)%>" size="5"> 
显示列数：<INPUT name="m106" type="text" value="<%=m_(106)%>" size="6">
显示翻页：<INPUT name="m107" type="checkbox" value="1" <%IF strint(m_(107))=1 Then%>checked<%End IF%>>

    </TR>

    <TR bgcolor="#ffffff">
      <TD height="26" align="right">推荐信息图文显示设置：<%mb="F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2">
	        信息类型：<INPUT name="m108" type="checkbox" value="1" <%IF strint(m_(108))=1 Then%>checked<%End IF%>>
			启用固顶：<INPUT name="m109" type="checkbox" value="1" <%IF strint(m_(109))=1 Then%>checked<%End IF%>>
			推荐标志：<INPUT name="m110" type="checkbox"  value="1" <%IF strint(m_(110))=1 Then%>checked<%End IF%>>
           	图片标志：<INPUT name="m111" type="checkbox"  value="1" <%IF strint(m_(111))=1 Then%>checked<%End IF%>>						
			点击数量：<INPUT name="m112" type="checkbox"  value="1" <%IF strint(m_(112))=1 Then%>checked<%End IF%>>
			显示日期(0-15)：<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">【例】</a>
			<INPUT name="m113" type="text" value="<%=m_(113)%>" size="5">
市场价格：<INPUT name="m114" type="checkbox"  value="1" <%IF strint(m_(114))=1 Then%>checked<%End IF%>>
本站价格：<INPUT name="m115" type="checkbox"  value="1" <%IF strint(m_(115))=1 Then%>checked<%End IF%>><br>
标题长度：<INPUT name="m116" type="text" value="<%=m_(116)%>" size="5">
标题颜色：<INPUT name="m117" type="checkbox" value="1" <%IF strint(m_(117))=1 Then%>checked<%End IF%>>
字体大小：<INPUT name="m118" type="text" value="<%=m_(118)%>" size="5">px
图片宽度：<INPUT name="m119" type="text" value="<%=m_(119)%>" size="6">
图片高度：<INPUT name="m120" type="text" value="<%=m_(120)%>" size="6">
单元宽度：<INPUT name="m121" type="text" value="<%=m_(121)%>" size="6">
单元高度：<INPUT name="m122" type="text" value="<%=m_(122)%>" size="6"><br>
每页个数：<INPUT name="m123" type="text" value="<%=m_(123)%>" size="5"> 
显示列数：<INPUT name="m124" type="text" value="<%=m_(124)%>" size="6">
显示翻页：<INPUT name="m125" type="checkbox" value="1" <%IF strint(m_(125))=1 Then%>checked<%End IF%>>

    </TR>
	
	    <TR  bgcolor="#BDDEEF">
      <TD height="26" align="right">热门信息图文显示设置：<%mb="F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2">
	        信息类型：<INPUT name="m126" type="checkbox" value="1" <%IF strint(m_(126))=1 Then%>checked<%End IF%>>
			启用固顶：<INPUT name="m127" type="checkbox" value="1" <%IF strint(m_(127))=1 Then%>checked<%End IF%>>
			推荐标志：<INPUT name="m128" type="checkbox"  value="1" <%IF strint(m_(128))=1 Then%>checked<%End IF%>>
           	图片标志：<INPUT name="m129" type="checkbox"  value="1" <%IF strint(m_(129))=1 Then%>checked<%End IF%>>						
			点击数量：<INPUT name="m130" type="checkbox"  value="1" <%IF strint(m_(130))=1 Then%>checked<%End IF%>>
			显示日期(0-15)：<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">【例】</a>
			<INPUT name="m131" type="text" value="<%=m_(131)%>" size="5">
市场价格：<INPUT name="m132" type="checkbox"  value="1" <%IF strint(m_(132))=1 Then%>checked<%End IF%>>
本站价格：<INPUT name="m133" type="checkbox"  value="1" <%IF strint(m_(133))=1 Then%>checked<%End IF%>><br>
标题长度：<INPUT name="m134" type="text" value="<%=m_(134)%>" size="5">
标题颜色：<INPUT name="m135" type="checkbox" value="1" <%IF strint(m_(135))=1 Then%>checked<%End IF%>>
字体大小：<INPUT name="m136" type="text" value="<%=m_(136)%>" size="5">px
图片宽度：<INPUT name="m137" type="text" value="<%=m_(137)%>" size="6">
图片高度：<INPUT name="m138" type="text" value="<%=m_(138)%>" size="6">
单元宽度：<INPUT name="m139" type="text" value="<%=m_(139)%>" size="6">
单元高度：<INPUT name="m140" type="text" value="<%=m_(140)%>" size="6"><br>
每页个数：<INPUT name="m141" type="text" value="<%=m_(141)%>" size="5"> 
显示列数：<INPUT name="m142" type="text" value="<%=m_(142)%>" size="6">
显示翻页：<INPUT name="m143" type="checkbox" value="1" <%IF strint(m_(143))=1 Then%>checked<%End IF%>>

    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">店铺显示设置：<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2">
显示个数：<INPUT name="m144" type="text" value="<%=m_(144)%>" size="5">
显示列数：<INPUT name="m145" type="text" value="<%=m_(145)%>" size="5">
标题长度：<INPUT name="m146" type="text" value="<%=m_(146)%>" size="5"> 
标题行高：<INPUT name="m147" type="text" value="<%=m_(147)%>" size="5">
标题字号：<INPUT name="m148" type="text" value="<%=m_(148)%>" size="5">
启用固顶：<INPUT name="m149" type="checkbox" value="1" <%IF strint(m_(149))=1 Then%>checked<%End IF%>><br>
仅显VIP会员的：<INPUT name="m150" type="checkbox" value="1" <%IF strint(m_(150))=1 Then%>checked<%End IF%>>
图文显示：
<SELECT name="m151">
<OPTION value="0" <%IF strint(m_(151))=0 Then%>selected<%End IF%>>文字</OPTION>
<OPTION value="1" <%IF strint(m_(151))=1 Then%>selected<%End IF%>>图文</OPTION>
</SELECT>
图片宽度：<INPUT name="m152" type="text" value="<%=m_(152)%>" size="5"> 
图片行高：<INPUT name="m153" type="text" value="<%=m_(153)%>" size="5">
</TD>
    </TR>
	   <TR  bgcolor="#BDDEEF">
      <TD height="26" align="right">店铺资源共享显示设置：<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2">
            显示栏目：<INPUT name="m154" type="text" value="<%=m_(154)%>" size="5">（填栏目ID号，0为全部）
			显示类型：
            <SELECT name="m155">
              <OPTION value="0" <%IF strint(m_(155))=0 Then%>selected<%End IF%>>全部</OPTION>
              <OPTION value="1" <%IF strint(m_(155))=1 Then%>selected<%End IF%>>推荐</OPTION>			  
			  <OPTION value="2" <%IF strint(m_(155))=2 Then%>selected<%End IF%>>热门</OPTION>			  
            </SELECT>
	  	    启用固顶：<INPUT name="m156" type="checkbox" value="1" <%IF strint(m_(156))=1 Then%>checked<%End IF%>>
	  	    推荐标志：<INPUT name="m157" type="checkbox" value="1" <%IF strint(m_(157))=1 Then%>checked<%End IF%>>
	  	    图片标志：<INPUT name="m158" type="checkbox" value="1" <%IF strint(m_(158))=1 Then%>checked<%End IF%>>
	  	    点击数：<INPUT name="m159" type="checkbox" value="1" <%IF strint(m_(159))=1 Then%>checked<%End IF%>>
			显示日期(0-15)：<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">【例】</a>
			<INPUT name="m160" type="text" value="<%=m_(160)%>" size="5"><br>
			标题长度：<INPUT name="m161" type="text" value="<%=m_(161)%>" size="5">
			显示条数：<INPUT name="m162" type="text" value="<%=m_(162)%>" size="5">
			显示列数：<INPUT name="m163" type="text" value="<%=m_(163)%>" size="5">
			标题字号：<INPUT name="m164" type="text" value="<%=m_(164)%>" size="5">
			标题行高：<INPUT name="m165" type="text" value="<%=m_(165)%>" size="5">
            显示更多：<INPUT name="m166" type="checkbox" value="1" <%IF strint(m_(166))=1 Then%>checked<%End IF%>>

</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">企业名片(文字)显示设置：<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2">
显示条数：<INPUT name="m167" type="text" value="<%=m_(167)%>" size="5">
显示列数：<INPUT name="m168" type="text" value="<%=m_(168)%>" size="5"> 
标题长度：<INPUT name="m169" type="text" value="<%=m_(169)%>" size="5">
标题行高：<INPUT name="m170" type="text" value="<%=m_(170)%>" size="5">
主营长度：<INPUT name="m171" type="text" value="<%=m_(171)%>" size="5">
地址长度：<INPUT name="m172" type="text" value="<%=m_(172)%>" size="5">
表格宽度：<INPUT name="m173" type="text" value="<%=m_(173)%>" size="5"><br>
标题加粗：<INPUT name="m174" type="checkbox" value="1" <%IF strint(m_(174))=1 Then%>checked<%End IF%>>
信笺背景：
            <SELECT name="m175">
              <OPTION value="0" <%IF strint(m_(175))=0 Then%>selected<%End IF%>>要信笺背景</OPTION>
              <OPTION value="1" <%IF strint(m_(175))=1 Then%>selected<%End IF%>>不要信笺保格式</OPTION>
			  <OPTION value="2" <%IF strint(m_(175))=2 Then%>selected<%End IF%>>不要信笺无格式</OPTION>			 
            </SELECT>
标题颜色：<INPUT name="m176" type="text" style="background:<%=m_(176)%>" onClick="Getcolor(this)" value="<%=m_(176)%>" size="8" maxlength="7">
显示主营地址电话：<INPUT name="m177" type="checkbox" value="1" <%IF strint(m_(177))=1 Then%>checked<%End IF%>>
显示更多：<INPUT name="m178" type="checkbox" value="1" <%IF strint(m_(178))=1 Then%>checked<%End IF%>>

</TD>
    </TR>

    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" bgcolor="#BDDEEF">企业名片(图片)显示设置：<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2" bgcolor="#BDDEEF" >
每行数量：<INPUT name="m179" type="text" value="<%=m_(179)%>" size="5">
显示总数：<INPUT name="m180" type="text" value="<%=m_(180)%>" size="5">
标题长度：<INPUT name="m181" type="text" value="<%=m_(181)%>" size="5"> 
主营长度：<INPUT name="m182" type="text" value="<%=m_(182)%>" size="5">

		</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" >新闻资讯一显示设置：<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%><br>（默认用于新闻动态）&nbsp;&nbsp;&nbsp;&nbsp;</TD>
      <TD colspan="2">
			显示类型：
            <SELECT name="m183">
              <OPTION value="0" <%IF strint(m_(183))=0 Then%>selected<%End IF%>>全部</OPTION>
              <OPTION value="1" <%IF strint(m_(183))=1 Then%>selected<%End IF%>>推荐</OPTION>
			  <OPTION value="2" <%IF strint(m_(183))=2 Then%>selected<%End IF%>>图文</OPTION>
			  <OPTION value="3" <%IF strint(m_(183))=3 Then%>selected<%End IF%>>热门</OPTION>			  
            </SELECT>
            一级栏目ID号：<INPUT name="m184" type="text" value="<%=m_(184)%>" size="3">
			二级栏目ID号：<INPUT name="m185" type="text" value="<%=m_(185)%>" size="3">
			三级栏目ID号：<INPUT name="m186" type="text" value="<%=m_(186)%>" size="3"><FONT color="#ff0000">（分类ID号在新闻资讯管理－文章栏目管理中获得，不能乱填，为0时显示全部）</font><br>
	  	    启用固顶：<INPUT name="m187" type="checkbox" value="1" <%IF strint(m_(187))=1 Then%>checked<%End IF%>>
	  	    推荐标志：<INPUT name="m188" type="checkbox" value="1" <%IF strint(m_(188))=1 Then%>checked<%End IF%>>
	  	    图片标志：<INPUT name="m189" type="checkbox" value="1" <%IF strint(m_(189))=1 Then%>checked<%End IF%>>
	  	    点击数：<INPUT name="m190" type="checkbox" value="1" <%IF strint(m_(190))=1 Then%>checked<%End IF%>>
	  	    滚动显示：<INPUT name="m191" type="checkbox" value="1" <%IF strint(m_(191))=1 Then%>checked<%End IF%>>
			显示日期(0-15)：<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">【例】</a>
			<INPUT name="m192" type="text" value="<%=m_(192)%>" size="3">			
			显示条数：<INPUT name="m193" type="text" value="<%=m_(193)%>" size="3">
			显示列数：<INPUT name="m194" type="text" value="<%=m_(194)%>" size="3"><br>
			标题长度：<INPUT name="m195" type="text" value="<%=m_(195)%>" size="3">
			标题字号：<INPUT name="m196" type="text" value="<%=m_(196)%>" size="3">
			标题行高：<INPUT name="m197" type="text" value="<%=m_(197)%>" size="3">

</TD>
    </TR>
	   <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" bgcolor="#BDDEEF">新闻资讯二显示设置：<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%><br>（默认用于服务指南）&nbsp;&nbsp;&nbsp;&nbsp;</TD>
      <TD colspan="2" bgcolor="#BDDEEF">
			显示类型：
            <SELECT name="m198">
              <OPTION value="0" <%IF strint(m_(198))=0 Then%>selected<%End IF%>>全部</OPTION>
              <OPTION value="1" <%IF strint(m_(198))=1 Then%>selected<%End IF%>>推荐</OPTION>
			  <OPTION value="2" <%IF strint(m_(198))=2 Then%>selected<%End IF%>>图文</OPTION>
			  <OPTION value="3" <%IF strint(m_(198))=3 Then%>selected<%End IF%>>热门</OPTION>			  
            </SELECT>
            一级栏目ID号：<INPUT name="m199" type="text" value="<%=m_(199)%>" size="3">
			二级栏目ID号：<INPUT name="m200" type="text" value="<%=m_(200)%>" size="3">
			三级栏目ID号：<INPUT name="m201" type="text" value="<%=m_(201)%>" size="3"><FONT color="#ff0000">（分类ID号在新闻资讯管理－文章栏目管理中获得，不能乱填，为0时显示全部）</font><br>
	  	    启用固顶：<INPUT name="m202" type="checkbox" value="1" <%IF strint(m_(202))=1 Then%>checked<%End IF%>>
	  	    推荐标志：<INPUT name="m203" type="checkbox" value="1" <%IF strint(m_(203))=1 Then%>checked<%End IF%>>
	  	    图片标志：<INPUT name="m204" type="checkbox" value="1" <%IF strint(m_(204))=1 Then%>checked<%End IF%>>
	  	    点击数：<INPUT name="m205" type="checkbox" value="1" <%IF strint(m_(205))=1 Then%>checked<%End IF%>>
	  	    滚动显示：<INPUT name="m206" type="checkbox" value="1" <%IF strint(m_(206))=1 Then%>checked<%End IF%>>
			显示日期(0-15)：<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">【例】</a>
			<INPUT name="m207" type="text" value="<%=m_(207)%>" size="3">
			显示条数：<INPUT name="m208" type="text" value="<%=m_(208)%>" size="3">
			显示列数：<INPUT name="m209" type="text" value="<%=m_(209)%>" size="3"><br>
			标题长度：<INPUT name="m210" type="text" value="<%=m_(210)%>" size="3">
			标题字号：<INPUT name="m211" type="text" value="<%=m_(211)%>" size="3">			
			标题行高：<INPUT name="m212" type="text" value="<%=m_(212)%>" size="3">

</TD>
    </TR>
   <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">新闻资讯三显示设置：<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2">
			显示类型：
            <SELECT name="m213">
              <OPTION value="0" <%IF strint(m_(213))=0 Then%>selected<%End IF%>>全部</OPTION>
              <OPTION value="1" <%IF strint(m_(213))=1 Then%>selected<%End IF%>>推荐</OPTION>
			  <OPTION value="2" <%IF strint(m_(213))=2 Then%>selected<%End IF%>>图文</OPTION>
			  <OPTION value="3" <%IF strint(m_(213))=3 Then%>selected<%End IF%>>热门</OPTION>			  
            </SELECT>
            一级栏目ID号：<INPUT name="m214" type="text" value="<%=m_(214)%>" size="3">
			二级栏目ID号：<INPUT name="m215" type="text" value="<%=m_(215)%>" size="3">
			三级栏目ID号：<INPUT name="m216" type="text" value="<%=m_(216)%>" size="3"><FONT color="#ff0000">（分类ID号在新闻资讯管理－文章栏目管理中获得，不能乱填，为0时显示全部）</font><br>
	  	    启用固顶：<INPUT name="m217" type="checkbox" value="1" <%IF strint(m_(217))=1 Then%>checked<%End IF%>>
	  	    推荐标志：<INPUT name="m218" type="checkbox" value="1" <%IF strint(m_(218))=1 Then%>checked<%End IF%>>
	  	    图片标志：<INPUT name="m219" type="checkbox" value="1" <%IF strint(m_(219))=1 Then%>checked<%End IF%>>
	  	    点击数：<INPUT name="m220" type="checkbox" value="1" <%IF strint(m_(220))=1 Then%>checked<%End IF%>>
	  	    滚动显示：<INPUT name="m221" type="checkbox" value="1" <%IF strint(m_(221))=1 Then%>checked<%End IF%>>
			显示日期(0-15)：<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">【例】</a>
			<INPUT name="m222" type="text" value="<%=m_(222)%>" size="3">
			显示条数：<INPUT name="m223" type="text" value="<%=m_(223)%>" size="3">
			显示列数：<INPUT name="m224" type="text" value="<%=m_(224)%>" size="3"><br>
			标题长度：<INPUT name="m225" type="text" value="<%=m_(225)%>" size="3">
			标题字号：<INPUT name="m226" type="text" value="<%=m_(226)%>" size="3">
			标题行高：<INPUT name="m227" type="text" value="<%=m_(227)%>" size="3">

</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" bgcolor="#BDDEEF">新闻资讯四显示设置：<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2" bgcolor="#BDDEEF">
	  		显示类型：
            <SELECT name="m228">
              <OPTION value="0" <%IF strint(m_(228))=0 Then%>selected<%End IF%>>全部</OPTION>
              <OPTION value="1" <%IF strint(m_(228))=1 Then%>selected<%End IF%>>推荐</OPTION>
			  <OPTION value="2" <%IF strint(m_(228))=2 Then%>selected<%End IF%>>图文</OPTION>
			  <OPTION value="3" <%IF strint(m_(228))=3 Then%>selected<%End IF%>>热门</OPTION>			  
            </SELECT>
            一级栏目ID号：<INPUT name="m229" type="text" value="<%=m_(229)%>" size="3">
			二级栏目ID号：<INPUT name="m230" type="text" value="<%=m_(230)%>" size="3">
			三级栏目ID号：<INPUT name="m231" type="text" value="<%=m_(231)%>" size="3"><FONT color="#ff0000">（分类ID号在新闻资讯管理－文章栏目管理中获得，不能乱填，为0时显示全部）</font><br>
	  	    启用固顶：<INPUT name="m232" type="checkbox" value="1" <%IF strint(m_(232))=1 Then%>checked<%End IF%>>
	  	    推荐标志：<INPUT name="m233" type="checkbox" value="1" <%IF strint(m_(233))=1 Then%>checked<%End IF%>>
	  	    图片标志：<INPUT name="m234" type="checkbox" value="1" <%IF strint(m_(234))=1 Then%>checked<%End IF%>>
	  	    点击数：<INPUT name="m235" type="checkbox" value="1" <%IF strint(m_(235))=1 Then%>checked<%End IF%>>
	  	    滚动显示：<INPUT name="m236" type="checkbox" value="1" <%IF strint(m_(236))=1 Then%>checked<%End IF%>>
			显示日期(0-15)：<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">【例】</a>
			<INPUT name="m237" type="text" value="<%=m_(237)%>" size="3">
			显示条数：<INPUT name="m238" type="text" value="<%=m_(238)%>" size="3">
			显示列数：<INPUT name="m239" type="text" value="<%=m_(239)%>" size="3"><br>			
			标题长度：<INPUT name="m240" type="text" value="<%=m_(240)%>" size="3">
			标题字号：<INPUT name="m241" type="text" value="<%=m_(241)%>" size="3">
			标题行高：<INPUT name="m242" type="text" value="<%=m_(242)%>" size="3">

</TD>
    </TR>	
	<TR bgcolor="#FFFFFF">
      <TD height="26" align="right">新闻资讯相册显示设置：<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2">
            一级栏目ID号：<INPUT name="m243" type="text" value="<%=m_(243)%>" size="3">
			二级栏目ID号：<INPUT name="m244" type="text" value="<%=m_(244)%>" size="3">
			三级栏目ID号：<INPUT name="m245" type="text" value="<%=m_(245)%>" size="3"><FONT color="#ff0000">（重要！分类ID号在新闻资讯管理－文章栏目管理中获得，不能乱填，为0时显示全部）</font><br>
	  	    启用固顶：<INPUT name="m246" type="checkbox" value="1" <%IF strint(m_(246))=1 Then%>checked<%End IF%>>
			显示推荐：<INPUT name="m247" type="checkbox" value="1" <%IF strint(m_(247))=1 Then%>checked<%End IF%>>
	  	    无图文章显示：<INPUT name="m248" type="checkbox" value="1" <%IF strint(m_(248))=1 Then%>checked<%End IF%>>
            标题显示：<INPUT name="m249" type="checkbox" value="1" <%IF strint(m_(249))=1 Then%>checked<%End IF%>>
			标题长度：<INPUT name="m250" type="text" value="<%=m_(250)%>" size="3">
			图片宽度：<INPUT name="m251" type="text" value="<%=m_(251)%>" size="3">
			图片高度：<INPUT name="m252" type="text" value="<%=m_(252)%>" size="3"><br>
			每行列数：<INPUT name="m253" type="text" value="<%=m_(253)%>" size="3">
			显示总数：<INPUT name="m254" type="text" value="<%=m_(254)%>" size="3">
			显示翻页：<INPUT name="m255" type="checkbox" value="1" <%IF strint(m_(255))=1 Then%>checked<%End IF%>>

</TD>
    </TR>
<TR  bgcolor="#BDDEEF">
      <TD height="26" align="right" >电子杂志设置：<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%> </TD>
      <TD colspan="2">
每行显示列数：<INPUT name="m256" type="text" value="<%=m_(256)%>" size="5">
总共显示个数：<INPUT name="m257" type="text" value="<%=m_(257)%>" size="5">
图片宽度：<INPUT name="m258" type="text" value="<%=m_(258)%>" size="5">
图片高度：<INPUT name="m259" type="text" value="<%=m_(259)%>" size="5">
标题长度：<INPUT name="m260" type="text" value="<%=m_(260)%>" size="5">
		</TD>
    </TR>
<TR bgcolor="#FFFFFF">
      <TD height="26" align="right" >都市114显示设置：<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%> </TD>
      <TD colspan="2">
每行显示条数：<INPUT name="m261" type="text" value="<%=m_(261)%>" size="5">
总共显示条数：<INPUT name="m262" type="text" value="<%=m_(262)%>" size="5">
		(<FONT color="#ff0000">列宽要在本模板方案－CSS风格中找到.articleline114sy项，修改该段width:的值</font>)</TD>
    </TR>
	<TR bgcolor="#FFFFFF"> 
      <TD height="26" align="right" bgcolor="#BDDEEF">便民服务显示设置：<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2"  bgcolor="#BDDEEF">
	  	    显示总数：<INPUT name="m263" type="text" value="<%=m_(263)%>" size="5">
	  	    每行个数：<INPUT name="m264" type="text" value="<%=m_(264)%>" size="5">
			小格宽度：<INPUT name="m265" type="text" value="<%=m_(265)%>" size="5">
	  	    行高：<INPUT name="m266" type="text" value="<%=m_(266)%>" size="5">
			字号：<INPUT name="m267" type="text" value="<%=m_(267)%>" size="5">
	  	    标题长度：<INPUT name="m268" type="text" value="<%=m_(268)%>" size="5"><br>
	  	    背景色：<INPUT name="m269" type="checkbox" value="1" <%IF strint(m_(269))=1 Then%>checked<%End IF%>>
			加粗：<INPUT name="m270" type="checkbox" value="1" <%IF strint(m_(270))=1 Then%>checked<%End IF%>>
            下划线：<INPUT name="m271" type="checkbox" value="1" <%IF strint(m_(271))=1 Then%>checked<%End IF%>>
            显示更多：<INPUT name="m272" type="checkbox" value="1" <%IF strint(m_(272))=1 Then%>checked<%End IF%>>
			</TD>
    </TR>

	<TR bgcolor="#FFFFFF"> 
      <TD height="26" align="right">友情链接显示设置：<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2">
            显示方式：
            <SELECT name="m273">
              <OPTION value="0" <%IF StrInt(m_(273))=0 Then%>selected<%End IF%>>图文分显</OPTION>
              <OPTION value="1" <%IF StrInt(m_(273))=1 Then%>selected<%End IF%>>综合显示</OPTION>
              <OPTION value="2" <%IF StrInt(m_(273))=2 Then%>selected<%End IF%>>仅显文字</OPTION>  
            </SELECT>
			站名长度：<INPUT name="m274" type="text" value="<%=m_(274)%>" size="5">
	  	    每行高度：<INPUT name="m275" type="text" value="<%=m_(275)%>" size="5">			
            每行条数：<INPUT name="m276" type="text" value="<%=m_(276)%>" size="5"><br>
			综合显示总数：<INPUT name="m277" type="text" value="<%=m_(277)%>" size="5">
			综合显示LOGO占行数：<INPUT name="m278" type="text" value="<%=m_(278)%>" size="5">			
			LOGO独显数量：<INPUT name="m279" type="text" value="<%=m_(279)%>" size="5">
			文字独显数量：<INPUT name="m280" type="text" value="<%=m_(280)%>" size="5">	
			</TD>
    </TR>

	<TR  bgcolor="#BDDEEF"> 
      <TD height="26" align="right">客户留言显示设置：<%mb=""
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2">
	  	    留言状态：<INPUT name="m281" type="checkbox" value="1" <%IF strint(m_(281))=1 Then%>checked<%End IF%>>
	  	    显示作者：<INPUT name="m282" type="checkbox" value="1" <%IF strint(m_(282))=1 Then%>checked<%End IF%>>			
	  	    显示日期(0-15)：<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">【例】</a>
			<INPUT name="m283" type="text" value="<%=m_(283)%>" size="5"><br>
			显示条数：<INPUT name="m284" type="text" value="<%=m_(284)%>" size="5">
            标题长度：<INPUT name="m285" type="text" value="<%=m_(285)%>" size="5">
            信笺背景：
            <SELECT name="m286">
              <OPTION value="0" <%IF StrInt(m_(286))=0 Then%>selected<%End IF%>>要信笺背景</OPTION>
              <OPTION value="1" <%IF StrInt(m_(286))=1 Then%>selected<%End IF%>>不要信笺保格式</OPTION>
			  <OPTION value="2" <%IF StrInt(m_(286))=2 Then%>selected<%End IF%>>不要信笺无格式</OPTION> 
            </SELECT> 	
			</TD>
    </TR>
   <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">本站公告显示设置：<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[√]"
	  Else
	  response.write "[<FONT color=#ff0000>×</font>]"
	  End if%></TD>
      <TD colspan="2">
	  	    启用固顶：<INPUT name="m287" type="checkbox" value="1" <%IF strint(m_(287))=1 Then%>checked<%End IF%>>
	  	    点击数：<INPUT name="m288" type="checkbox" value="1" <%IF strint(m_(288))=1 Then%>checked<%End IF%>>
	  	    滚动显示：<INPUT name="m289" type="checkbox" value="1" <%IF strint(m_(289))=1 Then%>checked<%End IF%>>
			显示日期(0-15)：<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">【例】</a>
			<INPUT name="m290" type="text" value="<%=m_(290)%>" size="3">
			显示条数：<INPUT name="m291" type="text" value="<%=m_(291)%>" size="3">
			显示列数：<INPUT name="m292" type="text" value="<%=m_(292)%>" size="3">
			标题长度：<INPUT name="m293" type="text" value="<%=m_(293)%>" size="3">
			标题字号：<INPUT name="m294" type="text" value="<%=m_(294)%>" size="3">
			标题行高：<INPUT name="m295" type="text" value="<%=m_(295)%>" size="3">
</TD>
    </TR>
<table>
<tr>
	<td> </td>
</tr>
</table>
	
	<TR bgcolor="#FFFFFF"> 
      <TD height="30" colspan="3" align=center>	  
      <input name="Submit" type="submit" id="Submit" value="确定">
	  <input name="Submit1" type="button" id="Submit1" value="返回" onClick="location='admin_Template_Management.asp?id=<%=ID%>'"/>&nbsp;&nbsp;&nbsp;&nbsp; <a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')"><b>0-15种显示日期</b></a> </TD>
      </TR>
  </FORM>
</TABLE>
      <BR></TD>
    </TR>
</TABLE>
</BODY>                    
                    
</HTML>                    
