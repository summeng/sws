<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("05")
Dim action,rsssc,urlpage,dweb
action=request.querystring("action")
rsssc=request.querystring("rsssc")
urlpage=request.querystring("urlpage")
	city_oneid=Request("city_one")
	city_twoid=Request("city_two")
	city_threeid=Request("city_three")
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

	type_oneid=Request("type_one")
	type_twoid=Request("type_two")
	type_threeid=Request("type_three")
   If type_oneid="" Then type_oneid=0
   If type_twoid="" Then type_twoid=0
   If type_threeid="" Then type_threeid=0
   If type_oneid=0 Then type_one=webname
   IF type_oneid=0 then
	tttt=""
	elseIF type_threeid>0 then
	tttt=" and type_oneid="&type_oneid&" and type_twoid="&type_twoid&" and type_threeid="&type_threeid
	elseIF type_twoid>0 then
	tttt=" and type_oneid="&type_oneid&" and type_twoid="&type_twoid
	Else
	tttt=" and type_oneid="&type_oneid
	End If
If action="Production" Then tttt=""
Call OpenConn
set rs=conn.execute("select * from DNJY_city where id>0 and twoid=0  order by indexid "&DN_OrderType&"")
if rs.eof or rs.bof then
response.write "没有地区分类，无法生成分站Rss！"
Call CloseRs(rs)
response.End
End if
	If city_oneid>0 Then city_one=Conn.Execute("select city from DNJY_city where twoid=0 and id="&city_oneid)(0)
	If city_oneid>0 Then dweb=Conn.Execute("select dname from DNJY_city where twoid=0 and id="&city_oneid)(0)
	If city_twoid>0 Then city_two=Conn.Execute("select city from DNJY_city where id="&city_oneid&" and threeid=0 and twoid="&city_twoid)(0)
	If city_twoid>0 Then dweb=Conn.Execute("select dname from DNJY_city where id="&city_oneid&" and threeid=0 and twoid="&city_twoid)(0)
	If city_threeid<>0 Then city_three=Conn.Execute("select city from DNJY_city where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid)(0)
	If city_threeid<>0 Then dweb=Conn.Execute("select dname from DNJY_city where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid)(0)
	If type_oneid>0 Then type_one=Conn.Execute("select name from DNJY_type where twoid=0 and id="&type_oneid)(0)
	If type_twoid<>0 Then type_two=Conn.Execute("select name from DNJY_type where id="&type_oneid&" and threeid=0 and twoid="&type_twoid)(0)
	If type_threeid<>0 Then type_three=Conn.Execute("select name from DNJY_type where id="&type_oneid&" and twoid="&type_twoid&" and threeid="&type_threeid)(0)
If shows1=1 Then cityweb=dweb
If action="Production" Or rsssc="yes" then
dim str,fso,objStream
Server.ScriptTimeout=50000
str = "<?xml version='1.0' encoding='UTF-8'?>" & vbcrlf
str = str & "<rss version=""2.0"">" & vbcrlf
str = str & "<channel>" & vbcrlf
str = str & "<title>" & city_one&city_two&city_three & " 供求信息 RSS feed</title>" & vbcrlf 
str = str & "<link>http://"&cityweb&"/"&strInstallDir&"</link>" & vbcrlf  
str = str & "<description>"&city_one&city_two&city_three&"供求信息</description>" & vbcrlf

set rs=server.createobject("ADODB.recordset")
sql = "select top 50 * from DNJY_info where yz=1 and fsohtm=1"&ttcc&""&tttt&""
if jiaoyi=0 Then sql=sql&" and zt="&jiaoyi
if overdue=0 Then sql=sql&" and dqsj>"&DN_Now&""
sql=sql&" order by fbsj "&DN_OrderType&""
rs.open sql,conn,1,1

if rs.Eof then
str = str & "<item>尚未有最新信息</item>" & vbcrlf
end if
Do while not rs.Eof 
str = str & "<item>"
If asphtm=1 then
str = str & "<link>http://"&cityweb&"/"&strInstallDir&"html/"&rs("id")&".html</link>" & vbcrlf
Else
str = str & "<link>http://"&cityweb&"/"&strInstallDir&"show.asp?id="&rs("id")&"</link>" & vbcrlf
end if
str = str & "<title>"&rs("biaoti")&"</title>" & vbcrlf
str = str & "<creator>"&rs("name")&"</creator>" & vbcrlf
str = str & "<PubDate>"&rs("fbsj")&"</PubDate>" & vbcrlf

If rs("Readinfo")=1 Then
str = str & "<description><![CDATA[会员权限阅读]]></description>" & vbcrlf
elseIf rs("Readinfo")=2 Then
str = str & "<description><![CDATA[VIP会员权限阅读]]></description>" & vbcrlf
Else
str = str & "<description><![CDATA["&replace(replace(left(rs("memo"),200),"<","&lt;"),">","&gt;")&"]]></description>" & vbcrlf
End If

str = str & "</item>" & vbcrlf
rs.MoveNext          
loop                   
set rs=nothing
conn.Close
set conn = Nothing

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

If rsssc="yes" And urlpage="info_add_save.asp" Then
response.Write "<script language=javascript>alert('添加成功，并同步生成"&city_one&city_two&city_three&"的RSS聚合服务页!');location.replace('info_add.asp');</script>"
response.End
End If

If rsssc="yes" And urlpage="info_edit_save.asp" Then
response.Write "<script language=javascript>alert('修改成功，并同步生成"&city_one&city_two&city_three&"的RSS聚合服务页!!');setTimeout(""location.replace('infolist.asp')"",0);</script>"
response.End
End If

If rsssc="yes" And urlpage="info_yz.asp" Then
Response.Write "<p align=""center"">修改成功，并同步生成"&city_one&city_two&city_three&"的RSS聚合服务页</p>"
Response.Write "<script language=javascript>setTimeout(""location.replace('infolist.asp?page="&request("t1")&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&dick="&request("dick")&"&city_one="&request("city_one")&"&city_two="&request("city_two")&"&city_three="&request("city_three")&"')"",200)</script>"
response.End
End If

Set objStream = Nothing
If Not Err Then
response.Write "<script language=javascript>alert('成功生成"&city_one&city_two&city_three&"的RSS聚合服务页!');location.replace('../xml/rss"&city_oneid&"_"&city_twoid&"_"&city_threeid&".xml');</script>"
Response.End
End If
End If
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#6699cc">
<TR>
    <TD height="20" bgcolor="#799AE1" align="center" colspan="13"><FONT color="#FFFFFF" style="font-size:14px">生 成 分 站 RSS 聚 合</FONT></TD>
 </TR>
  <FORM name="form" method="post" action="?action=Production" >
  
    <tr> 
      <td height="30" colspan="3" align="center">&nbsp;</td>
    </tr>

    <tr> 
      <td width="188" height="25"><font color="#FF0000">*</font>所属地区：</td>
      <td colspan="2">
	  <%Dim rsdq
	  set rsdq=server.createobject("adodb.recordset")
set rsdq=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")
%>
      <SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
		dim count:		count = 0
        do while not rsdq.eof 
        %>
subcat[<%=count%>] = new Array("<%=rsdq("city")%>","<%=rsdq("id")%>","<%=rsdq("twoid")%>");
        <%
        count = count + 1
        rsdq.movenext
        loop
        rsdq.close
		set rsdq=nothing
        %>
onecount=<%=count%>;
</SCRIPT>
<%
set rsdq=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid>0")
%>
 <SCRIPT language = "JavaScript">
var onecount4;
onecount4=0;
subcat4 = new Array();
        <%
		dim count4:count4 = 0
        do while not rsdq.eof 
        %>
subcat4[<%=count4%>] = new Array("<%=rsdq("city")%>","<%=rsdq("id")%>","<%=rsdq("twoid")%>","<%=rsdq("threeid")%>");
        <%
        count4 = count4 + 1
        rsdq.movenext
        loop
        rsdq.close
		set rsdq = nothing
        %>
onecount4=<%=count4%>;

function changelocation(locationid)
    {
    document.form.city_two.length = 0; 
	document.form.city_two.options[0] = new Option('选择城市','');
	document.form.city_three.length = 0; 
	document.form.city_three.options[0] = new Option('选择城市','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.form.city_two.options[document.form.city_two.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
	
	function changelocation4(locationid,locationid1)
    {
    document.form.city_three.length = 0; 
    document.form.city_three.options[0] = new Option('选择城市','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount4; i++)
        {
            if (subcat4[i][2] == locationid)
            { 
			if (subcat4[i][1] == locationid1)
			{
                document.form.city_three.options[document.form.city_three.length] = new Option(subcat4[i][0], subcat4[i][3]);
            }        
        }
       } 
    }            
              </SCRIPT>
      <SELECT name="city_one" size="1" id="select2" onChange="changelocation(document.form.city_one.options[document.form.city_one.selectedIndex].value)">
        <OPTION value="" selected>选择城市</OPTION>
        <%set rsdq=conn.execute("select * from DNJY_city where id>0 and twoid=0 order by indexid asc")
if rsdq.eof or rsdq.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsdq.eof
response.write "<option value='"&rsdq("id")&"'>"&rsdq("city")&"</option>"
    rsdq.movenext
    loop
end if
rsdq.close
set rsdq = Nothing

%>
      </SELECT>
      <SELECT name="city_two" id="city_two" onChange="changelocation4(document.form.city_two.options[document.form.city_two.selectedIndex].value,document.form.city_one.options[document.form.city_one.selectedIndex].value)">
    <OPTION value="" selected>选择城市</OPTION>
	</SELECT>
	     <SELECT name="city_three" id="city_three">
        <OPTION value="" selected>选择城市</OPTION>
    </SELECT> （<font color="#FF0000">不选择时将生成总站最新信息</font>)</td>
    </tr>
    <tr> 
      <td height="30" colspan="3" align="center"> 
        <input type="submit" name="Submit" value="提交" class="input" >
        
        <input type="reset" name="Submit2" value="重置" class="input"> 
 </td>
    </tr>
  </form>
  <tr> 
    <td height="30" colspan="3" align="center">&nbsp;</td>
  </tr>
</table>
<%Call CloseDB(conn)%>
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
}
-->
</style>