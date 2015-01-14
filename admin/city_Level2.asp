<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file="../Include/md5.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("02")
Call OpenConn
If left(web,4)="www." Then web=Mid(""&web&"",4)
dim action,idd,dcity,mname,twoid,dcityadmin,dcitypass,dbpath,ycity,fdel,iPage,adid,pic,sql1,sql2,sql3,two
twoid=strint(request.QueryString("twoid"))
id=strint(request.QueryString("id"))
idd=strint(request.QueryString("idd"))
action=HtmlEncode(request.querystring("action"))
dcity=HtmlEncode(request("city"))
dcityadmin=HtmlEncode(request("cityadmin"))
dcitypass=HtmlEncode(request.form("citypass"))
dname=LCase(HtmlEncode(Request.form("dname")))
if dcity="" and action<>"" then
	response.write "<script LANGUAGE='javascript'>alert('请认真填写！');history.go(-1);</script>"
	response.end
end If
WebSetting =""
For i=0 To 12
	WebSetting = WebSetting & HtmlEncode(Request("w"&i)) & "∮∮∮"
Next
WebSetting = Left(WebSetting,Len(WebSetting)-3)
dbpath="../"

set rs=server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_city where id="&id&" and twoid=0",conn,1,3
mname=rs("city")
Call CloseRs(rs)


select case action
case "add"
If dcity<>"" Then  
   set rs=server.CreateObject("adodb.recordset")
	rs.open "select city from DNJY_city where id="&id&" and threeid=0  and city='"&dcity&"' ",conn,1,1
	if not rs.eof Then
	response.Write "<script language=javascript>alert('此地区名：["&dcity&"]已经存在，请重新选择一个!');location.replace('javascript:history.go(-1)');</script>"	
	set rs=Nothing
    response.end	
	end If
end If	
If dname<>"" Then  
   set rs=server.CreateObject("adodb.recordset")
	rs.open "select dname from DNJY_city where dname='"&dname&"' ",conn,1,1
	if not rs.eof Then
	response.Write "<script language=javascript>alert('此域名：["&dname&"]已经存在，请重新选择一个!');location.replace('javascript:history.go(-1)');</script>"		
	set rs=Nothing
    response.end	
	end If
end If
If dcityadmin<>"" Then  
    set rs=server.CreateObject("adodb.recordset")
	sql="select a.cityadmin,b.username from DNJY_city a,DNJY_user b where a.cityadmin='"&dcityadmin&"' or b.username='"&dcityadmin&"'"
	rs.open sql,conn,1,1
	if not rs.eof Then
	response.Write "<script language=javascript>alert('此地区管理员用户名：["&dcityadmin&"]在地区管理员或前台会员中已经存在，请重新选择一个!');location.replace('javascript:history.go(-1)');</script>"	
	set rs=Nothing
    response.end	
	end If
end If 

If dcityadmin<>"" And dcitypass="" Then  
	response.write "<script LANGUAGE='javascript'>alert('填写地区管理员必须填写密码！');history.go(-1);</script>"
	response.end
end If 		
if dcity<>"" then 
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_city where id="&id&" and twoid<>0 and threeid=0 and city='"&dcity&"'",conn,1,3
if rs.eof then
rs.close
rs.open "select * from DNJY_city where id="&id&" order by twoid "&DN_OrderType&"",conn,1,3
two=rs("twoid")+1
rs.close
rs.open "DNJY_city",conn,1,3
rs.AddNew
rs("id")=id
rs("city")=dcity
rs("dname")=dname
rs("cityadmin")=dcityadmin
rs("citypass")=md5(dcitypass)
rs("indexid")=strint(request("indexid"))
rs("indexshow")=HtmlEncode(request("indexshow"))
Rs("WebSetting")=WebSetting
rs("twoid")=two
rs("threeid")=0
i=rs("i")
city_one=conn.Execute("Select city From DNJY_city Where id="&id&"")(0)
rs.Update
Call CloseRs(rs)
end if  
end If

If dcityadmin<>"" then
set rs = Server.CreateObject("ADODB.RecordSet")'同时注册前台会员
sql="select username,password,useryz,zcdata,vipdq,city_oneid,city_one,city_twoid,city_two,city_threeid,cid from DNJY_user"
rs.open sql,conn,1,3
rs.addnew
rs("username")=dcityadmin
rs("password")=md5(dcitypass)
rs("useryz")=1
rs("city_oneid")=id
rs("city_one")=city_one
rs("city_twoid")=two
rs("city_two")=dcity
rs("city_threeid")=0
rs("cid")=i
rs("zcdata")=now()
rs("vipdq")=now()
rs.update
Call CloseRs(rs)
End If

case "edit"
If dcity<>"" Then  
   set rs=server.CreateObject("adodb.recordset")
	rs.open "select city from DNJY_city where id="&id&" and city='"&dcity&"' and i<>"&idd&"",conn,1,1
	if not rs.eof Then
	response.Write "<script language=javascript>alert('此地区名：["&dcity&"]已经存在，请重新选择一个!');location.replace('javascript:history.go(-1)');</script>"		
	set rs=Nothing
    response.end	
	end If
end If	
If dname<>"" Then  
   set rs=server.CreateObject("adodb.recordset")
	rs.open "select dname from DNJY_city where dname='"&dname&"' and i<>"&idd&"",conn,1,1
	if not rs.eof Then
	response.Write "<script language=javascript>alert('此域名：["&dname&"]已经存在，请重新选择一个!');location.replace('javascript:history.go(-1)');</script>"	
	set rs=Nothing
    response.end	
	end If
    end If

If dcityadmin<>"" Then  
    set rs=server.CreateObject("adodb.recordset")
	sql="select a.cityadmin,b.username from DNJY_city a,DNJY_user b where (a.cityadmin='"&dcityadmin&"' or b.username='"&dcityadmin&"') and a.i<>"&idd&" and b.cid<>"&idd&" "
	rs.open sql,conn,1,1
	if not rs.eof Then
	response.Write "<script language=javascript>alert('此地区管理员用户名：["&dcityadmin&"]在地区管理员或前台会员中已经存在，请重新选择一个!');location.replace('javascript:history.go(-1)');</script>"	
	set rs=Nothing
    response.end	
	end If
end If 

If dcityadmin<>"" And dcitypass="" Then  
	response.write "<script LANGUAGE='javascript'>alert('填写地区管理员必须填写密码！');history.go(-1);</script>"
	response.end
end If
		
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_city where threeid=0 and twoid="&twoid&" and id="&id,conn,1,3
ycity=rs("city")
rs("city")=dcity
rs("dname")=dname
rs("cityadmin")=dcityadmin
rs("indexid")=strint(request("indexid"))
rs("indexshow")=HtmlEncode(request("indexshow"))
Rs("WebSetting")=WebSetting
if dcitypass<>rs("citypass") Or (dcitypass<>"" And IsNull(rs("citypass"))) then
rs("citypass")=md5(dcitypass)
end If
If dcityadmin="" Then rs("citypass")=""
rs("AppId")=HtmlEncode(Request.form("AppId"))
rs("AppKey")=HtmlEncode(Request.form("AppKey"))
rs("API_QQCallBack")=HtmlEncode(Request.form("API_QQCallBack"))
i=rs("i")
city_one=conn.Execute("Select city From DNJY_city Where id="&id&"")(0)
rs.Update
Call CloseRs(rs)
sql="update DNJY_info set city_two='"&dcity&"' where city_threeid=0 and city_twoid="&twoid&" and city_oneid="&id
conn.execute sql

set rs = Server.CreateObject("ADODB.RecordSet")'如果前台没此会员同时注册前台会员
sql="select username,password,useryz,zcdata,vipdq,city_oneid,city_one,city_twoid,city_two,city_threeid,cid from DNJY_user where username='"&dcityadmin&"'"
rs.open sql,conn,1,3
if (rs.eof or rs.bof) And dcityadmin<>"" then
rs.addnew
rs("username")=dcityadmin
rs("password")=md5(dcitypass)
rs("useryz")=1
rs("city_oneid")=id
rs("city_one")=city_one
rs("city_twoid")=twoid
rs("city_two")=dcity
rs("city_threeid")=0
rs("cid")=i
rs("zcdata")=now()
rs("vipdq")=now()
rs.update
End If
Call CloseRs(rs)

set rs = Server.CreateObject("ADODB.RecordSet")'如果前台有此会员同时修改前台会员
sql="select username,password,useryz,zcdata,vipdq,city_oneid,city_one,city_twoid,city_two,city_threeid,cid from DNJY_user where username='"&dcityadmin&"'"
rs.open sql,conn,1,3
if (Not rs.eof and Not rs.bof) And dcityadmin<>"" then
rs("username")=dcityadmin
rs("password")=md5(dcitypass)
rs("useryz")=1
rs("city_oneid")=id
rs("city_one")=city_one
rs("city_twoid")=twoid
rs("city_two")=dcity
rs("city_threeid")=0
rs("cid")=i
rs("zcdata")=now()
rs("vipdq")=now()
rs.update
End If
Call CloseRs(rs)

case "del"
set rs=server.CreateObject("adodb.recordset")
sql="select * from DNJY_city where id="&id&" and threeid>0 and twoid ="&twoid
rs.open sql,conn,3,2
if not rs.eof Then
Response.Write ("<script language=javascript>alert('此地区下级分类非空，请先删除下级分类才能删除本分类!');history.go(-1);</script>")
set rs=Nothing
response.End
End If

set rs=server.CreateObject("adodb.recordset")
conn.execute ("delete from DNJY_city where id="&id&" and twoid="&twoid&"")
'删除关联地区的信息及相关文件开始
sql="select * from DNJY_info where city_oneid="&id&" and city_twoid="&twoid
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,3,2
Set fdel = CreateObject("Scripting.FileSystemObject")	
if not rs.eof then
 For iPage = 1 To rs.recordcount
 adid=rs("id")
 pic=rs("tupian")
rs.delete
rs.update
rs.movenext  
   '''''''''''删除html开始''''''''''''''''
				 if fdel.fileExists(Server.MapPath("../html/"&trim(adid)&".html")) then
                 fdel.DeleteFile(Server.MapPath("../html/"&trim(adid)&".html"))
                 end if
				 if pic<>"0" And fdel.fileExists(Server.MapPath("../UploadFiles/infopic/"&pic&"")) then
				 fdel.DeleteFile(Server.MapPath("../UploadFiles/infopic/"&pic&""))
				 end if             

'''''''''''删除html结束'''''''''''''''
'删除相关回复开始
sql1="delete from [DNJY_reply] where xxid="&adid&" "
rs.open sql1,conn,1,3
'删除相关回复END
'删除相关举报开始
sql2="delete from [DNJY_Bad_info] where infoid="&adid&" "
rs.open sql2,conn,1,3
'删除相关举报END
'删除相关收藏开始
sql3="delete from [DNJY_favorites] where scid='"&adid&"' "
rs.open sql3,conn,1,3
'删除相关收藏END
if rs.eof then exit for
next
end if
rs.close
'删除关联地区的信息及相关文件END

'删除关联地区的新闻公告及相关文件开始
set rs=server.CreateObject("adodb.recordset")
sql="select * from DNJY_news where city_oneid="&id&" and city_twoid="&twoid
rs.open sql,conn,3,2
if not rs.eof then
 For iPage = 1 To rs.recordcount
 adid=rs("id")
 pic=rs("firstImageName")
rs.delete
rs.update

rs.movenext  
   '''''''''''删除html开始''''''''''''''''
   				 if fdel.fileExists(Server.MapPath("../news/"&trim(adid)&".html")) then
                 fdel.DeleteFile(Server.MapPath("../news/"&trim(adid)&".html"))
                 end if
				 if pic<>"" and lcase(left(pic,7))<>"http://" And fdel.fileExists(Server.MapPath("../"&pic&"")) then
				 fdel.DeleteFile(Server.MapPath("../"&pic&""))
				 end if           
'''''''''''删除html结束'''''''''''''''
if rs.eof then exit for
next
end If
rs.close
'删除关联地区的新闻公告及相关文件END

'删除关联地区的都市114开始
set rs=server.CreateObject("adodb.recordset")
sql="select * from DNJY_114 where city_oneid ="&id&" and city_twoid="&twoid 
rs.open sql,conn,3,2
if not rs.eof then
For iPage = 1 To rs.recordcount
rs.delete
rs.update
rs.movenext 
if rs.eof then exit for end if
next
end If
rs.close
'删除关联地区的都市114END

'删除关联地区的留言开始
set rs=server.CreateObject("adodb.recordset")
sql="select * from DNJY_gbook where city_oneid ="&id&" and city_twoid="&twoid
rs.open sql,conn,3,2
if not rs.eof then
For iPage = 1 To rs.recordcount
rs.delete
rs.update
rs.movenext 
if rs.eof then exit for end if
next
end If
rs.close
'删除关联地区的留言END
'删除关联地区的友情链接开始
set rs=server.CreateObject("adodb.recordset")
sql="select * from DNJY_link where city_oneid ="&id&" and city_twoid="&twoid
rs.open sql,conn,3,2
if not rs.eof then
For iPage = 1 To rs.recordcount
rs.delete
rs.update
rs.movenext 
if rs.eof then exit for end if
next
end If

'删除关联地区的合作网站END
set fdel=nothing
rs.close:set rs=nothing                                 
conn.close:set conn=nothing 
response.Redirect "?id="&id
end select
%>
<HTML>

<HEAD>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<META name="GENERATOR" content="Microsoft FrontPage 4.0">
<META name="ProgId" content="FrontPage.Editor.Document">
<TITLE>地区分类管理</TITLE>
<LINK rel="stylesheet" type="text/css" href="../css/style.css">
<script language="javascript">
<!--//
function show(v){
if(document.getElementById(v).style.display=='none')
document.getElementById(v).style.display='';
else
document.getElementById(v).style.display='none';
}			
//-->
</script>
</HEAD>
<BODY>

<TABLE width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
  <TR> 
    <TD height="20" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">二&nbsp;级&nbsp;地 区 分 类 管 理</FONT></TD>
  </TR>
  <TR> 
    <TD bgcolor="#FFFFFF"> &nbsp;&nbsp;[<A href="city_Level1.asp">首页</A>]-&gt;<%=mname%><BR>
	  <TABLE width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
        <TR align="center" bgcolor="#FFFFFF" height="20"> 
          <TD width="30">编号</TD>
		  <TD width="30">ID号</TD>
          <TD width="200">地区名称</TD>
          <TD width="40">显示顺序</TD>
          <TD width="50">首页显示</TD>
		  <TD width="35">域名绑定</TD>
          <TD width="100">地区管理员</TD>
          <TD width="100">地区管理密码<bR>（8位以下）</TD>
          <TD width="150">管理操作</TD>
        </TR>
        <%set rs=server.CreateObject("adodb.recordset")
		  rs.Open "select * from DNJY_city where id="&id&" and twoid<>0 and threeid=0 order by twoid",conn,1,1
		  dim follows
		  if rs.EOF and rs.BOF then
          response.write"<tr bgcolor=#FFFFFF><td colspan='9'><p align='center'><font color='red'>暂无地区分类！</font></td></tr></table><br>"
		  follows=0
		  else
		  do while not rs.EOF
				WebSetting = Empty
				If Rs("WebSetting")<>"" And Not IsNull(Rs("WebSetting")) Then
					WebSetting=Split(Rs("WebSetting"),"∮∮∮")
				Else
					ReDim WebSetting(12)
				End If
		  i=i+1
		  %>
        <FORM name="form1" method="post" action="?action=edit&id=<%=int(rs("id"))%>&twoid=<%=int(rs("twoid"))%>&idd=<%=rs("i")%>">
          <TR bgcolor="#FFFFFF" align="center"> 
            <TD><%=i%></TD>
			<TD><%=trim(rs("twoid"))%></TD>
            <TD>
              <INPUT name="city" type="text" id="city" value="<%=trim(rs("city"))%>" size="30">
              [<A href="city_Level3.asp?id=<%=rs("id")%>&twoid=<%=rs("twoid")%>">子目录</A>]</TD>
            <TD>
              <INPUT name="indexid" type="text" size="3" value="<%=trim(rs("indexid"))%>">
            </TD>
            <TD>
              <SELECT name="indexshow">
                <OPTION value="yes" <%if rs("indexshow")="yes" then%>selected<%end if%>>显示</OPTION>
                <OPTION value="no" <%if rs("indexshow")="no" then%>selected<%end if%>>隐藏</OPTION>
              </SELECT>
            </TD>
            <TD>
              <%if rs("dname")<>"" Then
              response.write"<font color=#ff0000>√</font>"
              Else
              response.write"×"
			  end if%>
            </TD>  			
            <TD>
              <INPUT name="cityadmin" type="text" id="cityadmin" value="<%=trim(rs("cityadmin"))%>" size="15">
            </TD>
            <TD>
              <INPUT name="citypass" type="password" id="citypass" value="<%=trim(rs("citypass"))%>" size="15">
            </TD>
            <TD>
              <INPUT type="submit" name="Submit" value="修 改">
              &nbsp; 
              <INPUT type="button" name="DEL" onClick="{if(confirm('确定要删除这个地区分类吗？\n将删除此分类下所有信息\n此操作不可以恢复！')){location.href='?city=<%=rs("city")%>&id=<%=rs("id")%>&twoid=<%=rs("twoid")%>&action=del';}return false;}" value="删除" >  
			  &nbsp;&nbsp;<input type="button" value="属 性" onClick="show('s<%=Rs("twoid")%>');">
            </TD>
          </TR>
          <TR bgcolor="#FFFFFF" align="center" id="s<%=Rs("twoid")%>" style="display:none">
           <TD colspan="9">
					 <table width="700" border="0" cellpadding="0" cellspacing="0" align="center">
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">分站LOGO名称：</TD>
							 <TD align="left" bgcolor="#C2D3FC"><font color="ff0000"><%=Rs("id")%>_<%=Rs("twoid")%>_<%=Rs("threeid")%>.jpg</font> （200×80，放在/UploadFiles/logo目录下）</TD>
					       </TR>
					       </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC"><font color="ff0000">特别注意：</font></TD>
							 <TD align="left" bgcolor="#C2D3FC" width="500"><font color="ff0000">以下黑色项目，只要填写其中一项，分站联系信息等将显示本分站的，因此分站属性中的内容建议填写完整，否则，分站联络信息显示残缺不全。如果全空则显示总站的。</font></TD>
					       </TR>
					 		<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">二级域名：</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="dname" type="text" id="dname" maxlength="50" value="<%=Rs("dname")%>"> 前不带"http://"，后不带"/" <font color='red'>（若不设置分站管理不要填写）</font></TD>
					       </TR>

							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">网站名称：</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input type="text" name="w0" type="text" id="w0" value="<%=WebSetting(0)%>"></TD>
						  </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">标题显示：</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input type="text" name="w1" type="text" id="w1" value="<%=WebSetting(1)%>"> 用于标题显示，可与网站名相同</TD>
						  </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">网站说明：</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="w2" type="text" id="w2" value="<%=WebSetting(2)%>"></TD>
						  </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">搜索关键词：</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="w3" type="text" id="w3" value="<%=WebSetting(3)%>"> 供搜索引擎搜索的关键内容，关键字用“,”号分隔</TD>						  </TR>

							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">邮件地址：</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="w4" type="text" id="w4" value="<%=WebSetting(4)%>"></TD>
						  </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">联系电话：</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="w5" type="text" id="w5" value="<%=WebSetting(5)%>"></TD>
						  </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">联系手机：</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="w6" type="text" id="w6" value="<%=WebSetting(6)%>"></TD>
						  </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">传真号码：</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="w7" type="text" id="w7" value="<%=WebSetting(7)%>"></TD>
						  </TR>
							<TR bgcolor="#FFFFFF" align="center">

							 <TD align="right" bgcolor="#C2D3FC">公司名称：</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="w12" type="text" id="w12" value="<%=WebSetting(12)%>"></TD>
						  </TR>							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">联系地址：</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="w8" type="text" id="w8" value="<%=WebSetting(8)%>"></TD>
						  </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">邮政编码：</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="w9" type="text" id="w9" value="<%=WebSetting(9)%>"></TD>
						  </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">在线QQ：</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="w10" type="text" id="w10" value="<%=WebSetting(10)%>"> 多个QQ号用西文逗号隔开，QQ号与下面的昵称一一对应</TD>
						  </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">QQ昵称：</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="w11" type="text" id="w11" value="<%=WebSetting(11)%>"> 多个昵称用西文逗号隔开，昵称与上面的QQ号一一对应</TD>
						  </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">＝＝＝＝＝＝＝＝</TD>
							 <TD align="left" bgcolor="#C2D3FC">＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝</font></TD>
					       </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC"></TD>
							 <TD align="left" bgcolor="#C2D3FC"><font color="0000ff">会员QQ快捷登录入口参数</font></TD>
					       </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC"><font color="0000ff">QQ登录AppID：</font></TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="AppId" type="password" size="40" id="AppId" maxlength="50" value="<%=trim(Rs("AppId"))%>"> <font color="red">opensns.qq.com 申请到的appid,<a href="http://connect.qq.com/" target="_blank">点此申请</a>。</font></TD>
					       </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC"><font color="0000ff">QQ登录AppKey：</font></TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="AppKey" type="password" size="40" id="AppKey" maxlength="50" value="<%=Rs("AppKey")%>"> <font color="red">opensns.qq.com 申请到的appkey。</font></TD>
					       </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC"><font color="0000ff">QQ登录后跳转的地址：</font></TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="API_QQCallBack" type="text" size="60" id="API_QQCallBack"  readonly maxlength="150" value="<%="http://"&Rs("dname")&"/"&strInstallDir&"api/qq/user.asp"%>"> <font class="tips">QQ登录成功后跳转的地址,不可改。</font></TD>
					       </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right"></TD>
							 <TD align="left" height="25"></font></TD>
					       </TR>
					 </table>				 
					 
					 </TD>
          </TR>
        </FORM>
        <%
		rs.MoveNext
          loop
          follows=rs.RecordCount
          end if
          Call CloseRs(rs)
		  Call CloseDB(conn)%>
      </TABLE>
	<BR></TD>
  </TR>
</TABLE>
<BR>
<TABLE width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
  <TR> 
    <TD height="20" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">添 加 二&nbsp;级&nbsp;地 区</FONT></TD>
  </TR>
  <TR> 
    <TD bgcolor="#FFFFFF">
	<BR>
	  <TABLE width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
        <TR align="center" bgcolor="#FFFFFF" height="20"> 
          <TD width="30">编号 </TD>
          <TD>地区名称</TD>
          <TD>显示顺序</TD>
          <TD>首页显示</TD>
          <TD>地区域名<br><font color='red'>（若不设置分站管理不要填写地区域名）</font></TD>
          <TD>地区管理员<br><font color='red'>（若不设置分站管理可不填写）</font></TD>
          <TD>地区管理密码<br><font color='red'>（若不设置分站管理可不填写）</font><bR>（8位以下）</TD>
          <TD>确定操作</TD>
        </TR>
        <FORM name="form1" method="post" action="?action=add&id=<%=id%>">
          <TR align="center" bgcolor="#FFFFFF"> 
            <TD><%=follows+1%></TD>
            <TD>
              <INPUT name="city" type="text" id="city" size="30">
            </TD>
            <TD>
              <INPUT name="indexid" type="text" value="0" size="3">
            </TD>
            <TD>
              <SELECT name="indexshow">
                <OPTION value="yes">显示</OPTION>
                <OPTION value="no" selected>隐藏</OPTION>
              </SELECT>
            </TD>
           <TD>
              <INPUT name="dname" type="text" size="20" value="">
            </TD>
            <TD>
              <INPUT name="cityadmin" type="text" id="cityadmin" size="15">
            </TD>
            <TD>
              <INPUT name="citypass" type="password" id="citypass" size="15">
            </TD>
            <TD>
              <INPUT type="submit" name="Submit3" value="添 加">
            </TD>
          </TR>
        </FORM>
      </TABLE>
	<BR></TD>
  </TR>
</TABLE>
</BODY> 
</HTML> 
