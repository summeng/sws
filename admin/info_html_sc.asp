<!--#include file="../conn.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("05")%>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
	margin-top: 16px;
}
-->
</style><body leftmargin="0">

<%if asphtm=0 then
response.write "<script language=JavaScript>" & chr(13) & "alert('系统已关闭生成静态页功能，无法生成静态页！');" &"window.location='info_html.asp'" & "</script>"
response.end
end If
Function Ceil(value)'有小数加1取整
Dim return,Cei2
return = int(value)
Cei2=value-return
if Cei2>0 Then
Ceil = return + 1
Else
Ceil=value+0'就是Ceil=value多一个+0 强调返回值为数字型
End If
End Function
dim type1,data,ss,fsohtm,dbpath,name,username,leixin,biaoti,fbsj,userip,memo,hfcs,dianhua,mobile,userqq,email,xxmpname,dizhi,youbian,biaotixs,scjiage,jiage,sdmba,usernameid,namea,xxtp,rsAll,Allrecord,Allbb,CreateType,BeginID,EndID,oneid,twoid,threeid,tjfd,page,j
Dim ServerURL,ServerURL1,aa,objfso,htmout,http,wcolor,ncolor,cksj
type1=Lcase(request("type"))
name=Trim(request("name"))
ID=Strint(Request.QueryString("ID"))
ss=Strint(Request.QueryString("s"))
CreateType=Strint(Request("CreateType"))
BeginID = strint(Trim(Request("BeginID")))
EndID = strint(Trim(Request("EndID")))

type_oneid=strint(Request("oneid"))
type_twoid=strint(Request("twoid"))
type_threeid=strint(Request("threeid"))
   If type_oneid="" Then type_oneid=0
   If type_twoid="" Then type_twoid=0
   If type_threeid="" Then type_threeid=0
	IF type_oneid=0 then
	tttt=""
	elseIF type_threeid>0 then
	tttt=" and type_oneid="&type_oneid&" and type_twoid="&type_twoid&" and type_threeid="&type_threeid
	elseIF type_twoid>0 then
	tttt=" and type_oneid="&type_oneid&" and type_twoid="&type_twoid
	Else
	tttt=" and type_oneid="&type_oneid
	End If

Allrecord=Strint(Request.QueryString("Allrecord"))
page=Strint(Request.QueryString("page"))
j=Strint(Request.QueryString("j"))
If Request.QueryString("j")="" Then j=1
If type1<>"xinxi" and type1<>"username" and type1<>"news" Then
	Response.Write "<script>alert('参数错误码，立即返回！');history.back();</script>"
	Response.End
End If
Server.ScriptTimeOut=30000
Allbb=request("Allbb")
tjfd=""
if CreateType=2 Then tjfd=tjfd&" and id between "&BeginID&" and "&EndID&""
if CreateType=4 Then tjfd=tjfd&""&tttt&""
if CreateType=5 Then tjfd=tjfd&" and fsohtm=0"
if jiaoyi=0 Then tjfd=tjfd&" and zt="&jiaoyi
if overdue=0 Then tjfd=tjfd&" and dqsj>"&DN_Now&""
if type1="username" Then tjfd=" and username='"&name&"'"
If Allbb<>"ok" Then
set rsAll = Server.CreateObject("ADODB.RecordSet")
set rsAll=conn.execute("select count(*) from DNJY_info where Readinfo=0 and yz=1"&tjfd&"")
session("Allrecord")=rsAll(0)
RsAll.Close:Set RsAll=Nothing
End if
If type1="xinxi" Or type1="username" Then
set rs = Server.CreateObject("ADODB.RecordSet")
Sql="Select Top 10 id,fsohtm,city_oneid,city_twoid,city_threeid From DNJY_info where Readinfo=0 and yz=1"&tjfd&" And id>="&ID
rs.open sql,conn,1,3
End If

If type1="news" Then Sql="Select Top 50 id From DNJY_News Where id>="&ID

If Not Rs.eof Then data=Rs.getrows
If Not IsArray(Data) Then
	If ID=0 Then
		Response.Write "没有任何符合生成条件的信息！"
	Else	    
		Response.Write "生成完成，共生成 <FONT color=""#FF0000"">"&session("j")-1&"</font> 条信息。"
		session("Allrecord")=""		
	End If
	Call CloseRs(rs)
    Call CloseDB(conn)
	Response.End
End If
If SS=0 Then
	Response.Write "共需生成 <FONT color=""#FF0000"">"&session("Allrecord")&"</font> 条信息，每页生成 <FONT color=""#FF0000"">10</font> 条信息，共需要分 <FONT color=""#FF0000"">"& Ceil(session("Allrecord")/10)&"</font> 页生成。目前生成第 <FONT color=""#FF0000"">1</font> 页。"
Else
    Response.Write "共需生成  <FONT color=""#FF0000"">"&session("Allrecord")&" </font> 条信息，每页生成 <FONT color=""#FF0000"">10</font> 条信息，共需要分 <FONT color=""#FF0000"">"&Ceil(session("Allrecord")/10)&"</font> 页生成。目前生成第 <FONT color=""#FF0000"">"&page+1&"</font> 页。</font><br>"
	Response.Write "已生成 <FONT color=""#FF0000"">"&Allrecord&"</font> 条信息...<br>"
End If
'对已生成记录标注已生成========
For i=0 to Ubound(data,2)
Conn.Execute("Update DNJY_info Set fsohtm=1 where id="&data(0,i)&"")
Next
Rs.Close:Set Rs=Nothing

'===================================

	ServerURL=CStr(Request.ServerVariables("SCRIPT_NAME"))
	n=InStrRev(ServerURL,"/") '从右边第一个字符起查找"_"的位置,n为返回值
	ServerURL=left(ServerURL, n-1)'显示从左边数第"n-1"个字符前面的字符,
	n=InStrRev(ServerURL,"/") '从右边第一个字符起查找"_"的位置,n为返回值
	ServerURL=left(ServerURL, n)'显示从左边数第"n"个字符前面的字符,
	ServerURL="http://"&Request.ServerVariables("Server_NAME")&ServerURL

Response.Write "免费版无此功能！"
Call CloseRs(rs)
Call CloseDB(conn)
Response.end

Server.ScriptTimeOut=60
'Response.Write "<script>location.href=""""";</script>"
Call CloseRs(rs)
Call CloseDB(conn)
%>
<script language=javascript>
function countDown(secs,surl){
 //alert(surl);
 tiao.innerText=secs;
 if(--secs>0){
     setTimeout("countDown("+secs+",'"+surl+"')",1000);
     }
 else{

     location.href=surl;
     }
 }
</script>
<br><FONT color="#FF0000"><span id="tiao">3</span>秒后自动生成下一页...</font><script language=javascript>countDown(3,'<%="?Id="&data(0,Ubound(Data,2))+1& "&type="&type1&"&s="&ss+i-1&"&Allrecord="&Allrecord+i&"&page="&page+1&"&j="&j&"&Allbb=ok&CreateType="&CreateType&"&BeginID="&id&"&EndID="&EndID&"&oneid="&type_oneid&"&twoid="&type_twoid&"&threeid="&type_threeid&""%>');</script>