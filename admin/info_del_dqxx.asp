<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
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
<%dim username,tupian,objFSO,fileExt,sql1,rs1,sql2,sql3
dim rss,str,fso,objStream,rs3
Server.ScriptTimeout=50000
username=request.cookies("dick")("username")
Call OpenConn
set rs3=server.createobject("adodb.recordset")
sql="select id,username,c,tupian,city_oneid,city_twoid,city_threeid from [DNJY_info] where "&DN_Now&">=dqsj"&ttcc&""
rs3.open sql,conn,1,1
'for i=1 to rs3.recordcount 
do while not rs3.eof
city_oneid=rs3("city_oneid")
city_twoid=rs3("city_twoid")
city_threeid=rs3("city_threeid")
id=rs3("id")
if rs3("c")=1 Then
   tupian=rs3("tupian")
   fileExt=lcase(right(tupian,4))
   if fileEXT=".gif" or fileEXT=".jpg" then
   call deltu()'同时删除相关图片
   end if
end If

conn.execute "UPDATE DNJY_user SET xxts=xxts-1 WHERE username='"&rs3("username")&"'" '同时向用户数据库减少一条信息记录
conn.execute "UPDATE DNJY_user SET jf=jf-"&jf_2*2&" WHERE username='"&rs3("username")&"'" '同时向用户数据库加倍减少积分


'===============为重新生成相关的rss取参数===========================
   If city_oneid="" Then city_oneid=0
   If city_twoid="" Then city_twoid=0
   If city_threeid="" Then city_threeid=0
   If city_oneid=0 Then city_one=webname
   IF city_oneid=0 then
	ttcc=""
	elseIF city_threeid>0 then
	ttcc=" and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid
	elseIF city_twoid>0 then
	ttcc=" and city_oneid="&city_oneid&" and city_twoid="&city_twoid
	Else
	ttcc=" and city_oneid="&city_oneid
	End If
'===============================================
set rs=server.createobject("adodb.recordset")
sql="delete from [DNJY_info] where id="&cstr(id)'删除信息
rs.open sql,conn,1,3
sql1="delete from [DNJY_favorites] where scid='"&cstr(id)&"' "'删除相关收藏
rs.open sql1,conn,1,3
sql2="delete from [DNJY_reply] where xxid="&cstr(id)&" "'删除相关回复
rs.open sql2,conn,1,3
sql3="delete from [DNJY_Bad_info] where infoid="&cstr(id)&" "'删除相关举报
rs.open sql3,conn,1,3

'================删除已生成的文件
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if objFSO.fileExists(Server.MapPath("../html/"&trim(id)&".html")) then
objFSO.DeleteFile(Server.MapPath("../html/"&trim(id)&".html"))
end if
set objfso=Nothing
'==================================================================
'===============重新生成相关的rss开始=====================================
str = "<?xml version='1.0' encoding='UTF-8'?>" & vbcrlf
str = str & "<rss version=""2.0"">" & vbcrlf
str = str & "<channel>" & vbcrlf
str = str & "<title />" & vbcrlf 
str = str & "<link>http://"&web&"</link>" & vbcrlf  
str = str & "<description>"&city_one&city_two&city_three&"供求信息</description>" & vbcrlf
set rss=server.createobject("ADODB.recordset")
sql = "select top 50 * from DNJY_info where yz=1 and fsohtm=1"&ttcc&""
if jiaoyi=0 Then sql=sql&" and zt="&jiaoyi
if overdue=0 Then sql=sql&" and dqsj>"&DN_Now&""
sql=sql&" order by fbsj "&DN_OrderType&""
rss.open sql,conn,1,1

if rss.Eof then
str = str & "<item>尚未有最新信息</item>" & vbcrlf
end if
Do while not rss.Eof 
str = str & "<item>"
If asphtm=1 then
str = str & "<link>http://"&web&"/"&strInstallDir&"html/"&rss("id")&".html</link>" & vbcrlf
Else
str = str & "<link>http://"&web&"/"&strInstallDir&"show.asp?id="&rss("id")&"</link>" & vbcrlf
End if
str = str & "<title>"&rss("biaoti")&"</title>" & vbcrlf
str = str & "<creator>"&rss("name")&"</creator>" & vbcrlf
str = str & "<PubDate>"&rss("fbsj")&"</PubDate>" & vbcrlf
str = str & "<description><![CDATA["&replace(replace(left(rss("memo"),200),"<","&lt;"),">","&gt;")&"]]></description>" & vbcrlf
str = str & "</item>" & vbcrlf
rss.MoveNext          
Loop
rss.close
set rss=nothing
str = str & "</channel>" & vbcrlf
str = str & "</rss>" & vbcrlf
set fso = nothing

Set objStream = Server.CreateObject("ADODB.Stream")
With objStream
.Open
.Charset = "UTF-8"
.Position = objStream.Size
.WriteText=str
.SaveToFile server.mappath("/"&strInstallDir&"xml/rss"&city_oneid&"_"&city_twoid&"_"&city_threeid&".xml"),2 '生成的XML文件名,不建议修改,如果你有更好的方案,可以改
.Close
End With
'===============重新生成相关的rss结束=====================================
'Next
i=i+1
if i>=rs3.recordcount then exit do 
rs3.movenext 
loop
Call CloseRs(rs3)

response.write "<script language=JavaScript>" & chr(13) & "alert('删除过期信息成功！');" &"window.location='infolist.asp?rsssc=yes&city_oneid="&city_oneid&"&city_twoid="&city_twoid&"&city_threeid&"&city_threeid&"&page="&request("page")&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&dick="&request("dick")&"&city_one="&request("city_one")&"&city_two="&request("city_two")&"&city_three="&request("city_three")&"'" & "</script>"
response.end


sub deltu()
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if objFSO.fileExists(Server.MapPath("../UploadFiles/infopic/"&tupian)) then
objFSO.DeleteFile(Server.MapPath("../UploadFiles/infopic/"&tupian))
end if
set objfso=nothing

end sub
Call CloseRs(rs)
Call CloseDB(conn)%>
