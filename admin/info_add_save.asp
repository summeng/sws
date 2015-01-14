<!--#include file="../conn.asp"-->
<!--#include file="../Include/err.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file="../Include/RemoteImg_save.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
%>
<%Dim webqq
webqq=qq
function chkdick(char) 
if instr(char,"'") then 
response.write "错误字符！"
response.end 
end if
if instr(char,"|") then
response.write "错误字符！"
response.end 
end if
if instr(char,"<") then
response.write "错误字符！"
response.end 
end if
if instr(char,">") then
response.write "错误字符！"
response.end 
end if
if instr(char,"&") then
response.write "错误字符！"
response.end 
end if
if instr(char,"%") then
response.write "错误字符！"
response.end 
end if
if instr(char,";") then
response.write "错误字符！"
response.end 
end if
if instr(char,"$") then
response.write "错误字符！"
response.end 
end if
end function
dim urlpage,biaoti,scjiage,jiage,memo,a,username,yz,CompPhone,youbian,dizhi,dqsj,sdays,hfje,fsohtm,homeEliteImage,gqje,userip
Dim leixin,fbsj,hfcs,dianhua,mobile,userqq,email,xxmpname,biaotixs,sdmba,usernameid,namea,xxtp,wcolor,ncolor,cksj,tpname,T_SaveImg,AspJpeg,T_UploadDir

city_oneid=strint(request.form("city_one"))
city_twoid=strint(request.form("city_two"))
city_threeid=strint(request.form("city_three"))
type_oneid=strint(request.form("type_one"))
type_twoid=strint(request.form("type_two"))
type_threeid=strint(request.form("type_three"))
Call OpenConn
set rs=server.createobject("adodb.recordset")
rs.open "select city from DNJY_city where twoid=0 and threeid=0 and id="&city_oneid,conn,1,1
city_one=rs("city")
rs.close
if city_twoid>0 then
rs.open "select city from DNJY_city where id="&city_oneid&" and threeid=0 and twoid="&city_twoid,conn,1,1
city_two=rs("city")
rs.close
end if
if city_twoid>0 and city_threeid>0 then
rs.open "select city from DNJY_city where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid,conn,1,1
city_three=rs("city")
rs.close
end if
rs.open "select name from DNJY_type where twoid=0 and threeid=0 and id="&type_oneid,conn,1,1
type_one=rs("name")
rs.close
if type_twoid>0 then
rs.open "select name from DNJY_type where id="&type_oneid&" and threeid=0 and twoid="&type_twoid,conn,1,1
type_two=rs("name")
rs.close
end if
if type_twoid>0 and type_threeid>0 then
rs.open "select name from DNJY_type where id="&type_oneid&" and threeid="&type_threeid&" and twoid="&type_twoid,conn,1,1
type_three=rs("name")
rs.close
end If
username=request.cookies("dick")("username")
biaoti=HTMLEncode(trim(request("biaoti")))
keywords=HTMLEncode(Replace_Text(trim(request.Form("keywords"))))
scjiage=trim(request("scjiage"))
jiage=trim(request("jiage"))
a=trim(request("a"))
yz=request("yz")
sdays=trim(request("days"))
If trim(request("memo"))="" Then
Call Alert ("请填写信息内容","-1")
ElseIf CheckStringLength(trim(request("memo")))>memonum Then
Call Alert ("信息内容限制在"&memonum&"个字符内","-1")
else
memo=trim(request.Form("memo"))
End If
tpname=request.form("tpname")



set rs=server.createobject("adodb.recordset")
sql = "select * from DNJY_info"
rs.open sql,conn,1,3
rs.addnew
rs("city_oneid")=city_oneid
rs("city_one")=city_one
rs("city_two")=city_two
rs("city_twoid")=city_twoid
rs("city_three")=city_three
rs("city_threeid")=city_threeid
rs("type_oneid")=type_oneid
rs("type_one")=type_one
rs("type_twoid")=type_twoid
rs("type_two")=type_two
rs("type_threeid")=type_threeid
rs("type_three")=type_three
rs("biaoti")=biaoti
rs("keywords")=keywords
rs("leixing")=request("leixing")
If scjiage<>"" Then rs("scjiage")=scjiage
rs("jiage")=jiage
rs("Readinfo")=trim(request.form("Readinfo"))
rs("memo")=memo
If tpname<>"" then
rs("tupian")=tpname
rs("c")=1
Else
rs("tupian")=0
rs("c")=0
End if
rs("a")=trim(request("a"))
rs("b")=trim(request("b"))
rs("username")=request.cookies("administrator")
rs("name")=HTMLEncode(trim(request("name")))
rs("email")=HTMLEncode(trim(request("email")))
rs("dianhua")=HTMLEncode(trim(request("dianhua")))
rs("CompPhone")=trim(request("CompPhone"))
If trim(request("qq"))<>"" Then rs("qq")=trim(request("qq"))
rs("youbian")=trim(request("youbian"))
rs("dizhi")=HTMLEncode(trim(request("dizhi")))
rs("weixin")=trim(request("weixin"))
rs("fax")=trim(request("fax"))
rs("mpname")=trim(request("mpname"))
rs("wcolor")=trim(request("wcolor"))
rs("ncolor")=trim(request("ncolor"))
rs("homeEliteImage")=trim(request("homeEliteImage"))
rs("hfje")=CInt(trim(request("hfje")))
If trim(request("hfje"))>0 then
rs("hfxx")=1
Else
rs("hfxx")=0
End If
rs("addsj")=now()
rs("dqsj")= dateadd("d", sdays, now)
rs("fbts")=sdays
rs("fbsj")=now()
rs("yz")=request("yz")
If request("yz")=1 then
rs("fsohtm")=1
Else
rs("fsohtm")=0
End If
userip=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If userip="" Then 
userip=Request.ServerVariables("REMOTE_ADDR")
end if
rs("ip")=userip
rs.update
Call CloseRs(rs)
set rs=server.createobject("ADODB.recordset")
sql = "select top 1 id from DNJY_info order by id "&DN_OrderType&""
rs.open sql,conn,1,1
id=rs("id")
Call CloseRs(rs)


'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝


		
'如果是竞价信息则做财务处理
If trim(request("hfje"))>0 Then
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_Financial] "
rs.open sql,conn,1,3
rs.addnew
rs("username")=request.cookies("administrator")
rs("aliname")=trim(request("name"))
rs("ywje")=trim(request("hfje"))
rs("ywlx")="支出"
rs("czbz")="发布竞价信息扣款"
rs("czz")=username
rs("admincl")=1
rs("addTime")=now()
rs.update
Call CloseRs(rs)
End IF
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
Call CloseDB(conn)
response.write "<p align=""center"">添加<font color=ff0000>"&biaoti&"</font>成功！</p>"
If request("yz")=1 then
urlpage="info_add_save.asp"
response.Write "<script language=javascript>location.replace('Substationrss.asp?rsssc=yes&urlpage="&urlpage&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&"&type_one="&type_oneid&"&type_two="&type_twoid&"&type_three="&type_threeid&"');</script>"
End IF%>
