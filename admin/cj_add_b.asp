<!--#include file="../conn.asp"-->
<!--#include file=cookies.asp-->
<!--#include file=../config.asp-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file="inc/Article_Function.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
response.buffer=true	
%>
<html>
<head>
<title>采集系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
</head>
<%Call OpenConn
dim FoundErr,ErrMsg,SqlItem,RsItem,ItemID,ItemName,WebUrl,ChannelID,strChannelDir,ItemDemo,LoginType,LoginUrl,LoginPostUrl,LoginUser,LoginPass,LoginFalse,tClass,tSpecial,c_oneid,c_twoid,c_threeid,c_one,c_two,c_three

ItemName=Trim(Request.Form("ItemName"))
WebName=Trim(Request.Form("WebName"))
WebUrl=Trim(Request.Form("WebUrl"))
city_oneid=strint(Trim(Request.Form("city_one")))
city_twoid=strint(Trim(Request.Form("city_two")))
city_threeid=strint(Trim(Request.Form("city_three")))
ItemDemo=Trim(Request.Form("ItemDemo"))
LoginType=Request.Form("LoginType")
LoginUrl=Trim(Request.Form("LoginUrl"))
LoginPostUrl=Trim(Request.Form("LoginPostUrl"))
LoginUser=Trim(Request.Form("LoginUser"))
LoginPass=Trim(Request.Form("LoginPass"))
LoginFalse=Trim(Request.Form("LoginFalse"))

If ItemName="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>项目名称不能为空</li>"
End If
If WebName="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>网站名称不能为空</li>"
End If
If WebUrl="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>网站网址不能为空</li>" 
End If

If city_oneid=0  Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>未指定地区</li>"
Else
   city_one=conn.execute("select city From DNJY_city Where id="  & city_oneid&" and twoid=0")(0)
   if city_twoid>0 Then city_two=conn.execute("select city From DNJY_city Where id="  & city_oneid&" and twoid="&city_twoid&" and threeid=0")(0)
   if city_twoid>0 and city_threeid>0 Then city_three=conn.execute("select city From DNJY_city Where id="  & city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid)(0)
End if

If request.form("c_oneid")="" Or request.form("c_oneid")=0  Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>未指定栏目</li>"
Else
'得出栏目名
c_oneid=strint(request.form("c_oneid"))
c_twoid=strint(request.form("c_twoid"))
c_threeid=strint(request.form("c_threeid"))
if c_oneid>0 then
set rs=server.createobject("adodb.recordset")
rs.open "select name from DNJY_news_class where twoid=0 and threeid=0 and oneid="&c_oneid,conn,1,1
c_one=rs("name")
rs.close
if c_twoid>0 then
rs.open "select name from DNJY_news_class where oneid="&c_oneid&" and threeid=0 and twoid="&c_twoid,conn,1,1
c_two=rs("name")
rs.close
end if
if c_twoid>0 and c_threeid>0 then
rs.open "select name from DNJY_news_class where oneid="&c_oneid&" and twoid="&c_twoid&" and threeid="&c_threeid,conn,1,1
c_three=rs("name")
rs.close
end if
set rs=nothing
end If
End if

If LoginType="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>请选择登录类型</li>"
Else
   LoginType=Clng(LoginType)
   If LoginType=1 Then
         If LoginUrl="" or LoginPostUrl="" or LoginUser="" or  LoginPass="" or  LoginFalse="" then
         FoundErr=True
         ErrMsg=ErrMsg& "<br><li>请将登录参数填写完整</li>"
      End If
   End If
End If

If FoundErr<>True Then
   SqlItem="Select * from DNJY_Article_Item"
   Set RsItem=server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,conn,1,3
   RsItem.AddNew
   RsItem("ItemName")=ItemName
   RsItem("WebName")=WebName
   RsItem("WebUrl")=WebUrl
	rsitem("city_oneid") = city_oneid
	rsitem("city_twoid")=city_twoid
	rsitem("city_threeid") = city_threeid
	rsitem("city_one") = city_one
	rsitem("city_two") = city_two
	rsitem("city_three") = city_three
    rsitem("c_oneid")=c_oneid
    rsitem("c_one")=c_one
    rsitem("c_twoid")=c_twoid
    rsitem("c_two")=c_two
    rsitem("c_threeid")=c_threeid
    rsitem("c_three")=c_three
	rsitem("mailType") =0 		
   If ItemDemo<>"" then
      RsItem("ItemDemo")=ItemDemo
   End if
   RsItem("LoginType")=LoginType
   If LoginType=1 Then
      RsItem("LoginUrl")=LoginUrl
      RsItem("LoginPostUrl")=LoginPostUrl
      RsItem("LoginUser")=LoginUser
      RsItem("LoginPass")=LoginPass
      RsItem("LoginFalse")=LoginFalse
   End If
   RsItem.UpDate
   RsItem.Close
   Set RsItem=Nothing
   SqlItem="Select top 1 Max(ItemID) from DNJY_Article_Item"
   ItemID=Conn.Execute(SqlItem)(0)
End If

If FoundErr=True Then
   call WriteErrMsg(ErrMsg)
Else
   Call Main
End If
'关闭数据库链接
Call CloseDB(conn)
%>
<%Sub Main%>
<HTML>
<HEAD>
<TITLE>新闻采集系统</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK rel="stylesheet" type="text/css" href="Admin_Style.css">
</HEAD>
<BODY leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<TABLE width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <TR> 
    <TD height="22" colspan="2" align="center" class="topbg"><strong>新闻采集系统项目管理</TD>
  </TR>
  <TR class="tdbg"> 
    <TD width="65" height="30"><STRONG>管理导航：</STRONG></TD>
    <TD height="30"><A href="cj_add_a.asp">添加项目</A> >> <A href="cj_editor_a.asp">基本设置</A> >> <FONT color=red>列表设置</FONT> >> 链接设置 >> 正文设置 >> 采样测试 >> 属性设置 >> 完成</TD>         
  </TR>         
</TABLE>
<BR>         
<TABLE width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
<FORM method="post" action="cj_add_c.asp" name="form1">
    <TR> 
      <TD height="22" colspan="2" class="title"> <DIV align="center"><STRONG>添 加 新 
		项 目--列  表 设 置</STRONG></DIV></TD>
    </TR>
    <TR class="tdbg"> 
      <TD width="20%" class="tdbg" align="right"><STRONG>列表索引页面：</STRONG></TD>
      <TD class="tdbg" width="75%">
		<INPUT name="ListStr" type="text" size="58" maxlength="200"> 
      </TD>
    </TR>
    <TR class="tdbg"> 
      <TD width="20%" class="tdbg" align="right" ><STRONG>列表开始标记：</STRONG></TD>
      <TD class="tdbg" width="75%">
      <TEXTAREA name="LsString" cols="49" rows="7"></TEXTAREA><BR>
      </TD>
    </TR>
    <TR class="tdbg"> 
      <TD width="20%" class="tdbg" align="right" ><STRONG>列表结束标记：</STRONG></TD>
      <TD class="tdbg" width="75%">
      <TEXTAREA name="LoString" cols="49" rows="7"></TEXTAREA><BR>
      </TD>
    </TR>
    <TR>
      <TD width="20%" class="tdbg" align="right"><STRONG> 列表索引分页：</STRONG></TD>
      <TD class="tdbg" width="75%">
		<INPUT type="radio" value="0" name="ListPaingType" checked onClick="ListPaing1.style.display='none';ListPaing12.style.display='none';ListPaing2.style.display='none';ListPaing3.style.display='none'">不作设置&nbsp;
		<INPUT type="radio" value="1" name="ListPaingType" onClick="ListPaing1.style.display='';ListPaing12.style.display='';ListPaing2.style.display='none';ListPaing3.style.display='none'">设置标签&nbsp;
		<INPUT type="radio" value="2" name="ListPaingType" onClick="ListPaing1.style.display='none';ListPaing12.style.display='none';ListPaing2.style.display='';ListPaing3.style.display='none'">批量生成&nbsp;
		<INPUT type="radio" value="3" name="ListPaingType" onClick="ListPaing1.style.display='none';ListPaing12.style.display='none';ListPaing2.style.display='none';ListPaing3.style.display=''">手动添加
	   </TD>
    </TR>
	<TR class="tdbg" id="ListPaing1" style="display:none">
      <TD width="20%" class="tdbg" align="right"><STRONG><FONT color=blue>下页开始标记：</FONT></STRONG><P>　</P><P>　</P>
      <STRONG><FONT color=blue>下页结束标记：</FONT></STRONG>
      </TD>
      <TD class="tdbg" width="75%">
		<TEXTAREA name="LPsString" cols="49" rows="7"></TEXTAREA><BR>
		<TEXTAREA name="LPoString" cols="49" rows="7"></TEXTAREA>
	   </TD>
    </TR>
	<TR class="tdbg" id="ListPaing12" style="display:none">
      <TD width="20%" class="tdbg" align="right"><STRONG><FONT color=blue>索引分页重定向：</FONT></STRONG>
      </TD>
      <TD class="tdbg" width="75%">
	<INPUT name="ListPaingStr1" type="text" size="58" maxlength="200" value=""><BR>
        格式：http://www.scuta.net/list.asp?page={$ID}  
	   </TD>
    </TR>
    <TR class="tdbg" id="ListPaing2" style="display:none"> 
      <TD width="20%" class="tdbg" align="right"><STRONG><FONT color=blue>批量生成：</FONT></STRONG></TD>
      <TD class="tdbg" width="75%">
        原字符串：<BR>
		<INPUT name="ListPaingStr2" type="text" size="58" maxlength="200" value=""><BR>
                格式：http://www.scuta.net/list.asp?page={$ID}<BR><BR>
	    生成范围：<BR>
		<INPUT name="ListPaingID1" type="text" size="8" maxlength="200"><SPAN lang="en-us"> To </SPAN><INPUT name="ListPaingID2" type="text" size="8" maxlength="200"><BR>
                格式：只能是数字，可升序或者降序。
      </TD>
    </TR>
    <TR class="tdbg" id="ListPaing3" style="display:none"> 
      <TD width="20%" class="tdbg" align="right"><STRONG><FONT color=blue>手动添加：</FONT></STRONG>
      </TD>
      <TD class="tdbg" width="75%">
      <TEXTAREA name="ListPaingStr3" cols="49" rows="7"></TEXTAREA><BR>
      格式：输入一个网址后按回车，再输入下一个。
       </TD>
    </TR>
    <TR class="tdbg"> 
      <TD colspan="2" align="center" class="tdbg">
        <INPUT name="ItemID" type="hidden" value="<%=ItemID%>">
        <INPUT  type="submit" name="Submit" value="下&nbsp;一&nbsp;步" style="cursor: hand;background-color: #cccccc;"></TD>
    </TR>
</FORM>
</TABLE>      

</BODY>         
</HTML>
<%End Sub%>
