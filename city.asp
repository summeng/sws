<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<%'作用：更改数据库中原有地区分类为新分类
Call OpenConn
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_user] where city_oneid=44 and city_twoid=3 and city_threeid=6" 
rs.open sql,conn,1,3
do while not rs.eof
rs("city_oneid")=44
rs("city_one")="广东"
rs("city_twoid")=3
rs("city_two")="深圳市"
rs("city_threeid")=4
rs("city_three")="福田区"
rs.update
Rs.movenext
Loop
Call CloseRs(rs)%>