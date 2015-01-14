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
Call OpenConn
Dim action,Rankrs,Rank,dweb
action=request.querystring("action")
Rank=Request("Rank")
If action="Production" Then
dim str,fso,objStream
Server.ScriptTimeout=50000
i=0
set Rankrs=server.createobject("adodb.recordset")
If Rank="1" then
sql = "select * from DNJY_city where id>0 and twoid=0 and threeid=0"
ElseIf Rank="2" then
sql = "select * from DNJY_city where id>0 and twoid>0 and threeid=0"
Else
sql = "select * from DNJY_city where id>0 and twoid>0 and threeid>0"
End if
Rankrs.open sql,conn,1,1
Do while not Rankrs.Eof
If Rank="1" then
ttcc=" and city_oneid="&Rankrs("id")&""
city_twoid=0
city_threeid=0
ElseIf Rank="2" then
ttcc=" and city_oneid="&Rankrs("id")&" and city_twoid="&Rankrs("twoid")&""
city_twoid=Rankrs("twoid")
city_threeid=0
Else
ttcc=" and city_oneid="&Rankrs("id")&" and city_twoid="&Rankrs("twoid")&" and city_threeid="&Rankrs("threeid")&""
city_twoid=Rankrs("twoid")
city_threeid=Rankrs("threeid")
End If
dweb=Rankrs("dname")
If shows1=1 Then cityweb=dweb
str = "<?xml version='1.0' encoding='UTF-8'?>" & vbcrlf
str = str & "<rss version=""2.0"">" & vbcrlf
str = str & "<channel>" & vbcrlf
str = str & "<title />" & vbcrlf 
str = str & "<link>http://"&cityweb&"/"&strInstallDir&"</link>" & vbcrlf  
str = str & "<description>"&Rankrs("city")&"供求信息</description>" & vbcrlf
set rs=server.createobject("ADODB.recordset")
sql = "select top 50 * from DNJY_info where yz=1 and fsohtm=1"&ttcc&""
if jiaoyi=0 Then sql=sql&" and zt="&jiaoyi
if overdue=0 Then sql=sql&" and dqsj>"&DN_Now&""
sql=sql&" order by fbsj "&DN_OrderType&""
rs.open sql,conn,1,1

if rs.Eof then
str = str & "<item>尚未有"&Rankrs("city")&"最新信息</item>" & vbcrlf
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


str = str & "</channel>" & vbcrlf
str = str & "</rss>" & vbcrlf
set fso = nothing

Set objStream = Server.CreateObject("ADODB.Stream")
With objStream
.Open
.Charset = "UTF-8"
.Position = objStream.Size
.WriteText=str
.SaveToFile server.mappath("/"&strInstallDir&"xml/rss"&Rankrs("id")&"_"&city_twoid&"_"&city_threeid&".xml"),2 '生成的XML文件名,不建议修改,如果你有更好的方案,可以改
.Close
End With
Set objStream = Nothing
i=i+1
Rankrs.MoveNext          
loop  

If Not Err Then
response.Write "成功批量生成RSS聚合服务页!共生成"&i&"页"
Response.End
End If
Rankrs.close
set Rankrs=Nothing
conn.Close
set conn = Nothing
End If
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#6699cc">
<TR>
    <TD height="20" bgcolor="#799AE1" align="center" colspan="13"><FONT color="#FFFFFF" style="font-size:14px">批 量 生 成 分 站 RSS 聚 合</FONT></TD>
 </TR>
  <FORM name="form" method="post" action="?action=Production" >
  
    <tr> 
      <td height="30" colspan="3" align="center">&nbsp;</td>
    </tr>

    <tr> 
      <td width="188" height="25">               <tr>
                  <td width="20%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF">
				  <p style="margin-left: 20">选择地区分类级别：<br>（分一级－二级－三级地区）
				
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20">
                <input type="radio" name="Rank" value="1" id="Rank"  checked>
                 全部一级 (30多页）&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="Rank" value="2" id="Rank">
                全部二级 （400多页）&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="radio" name="Rank" value="3" id="Rank">
                全部三级  （2500多页）<p>（<font color="#FF0000">生成三级时将占用大量服务器资源，并生成二千多页，占用近20M空间，请慎用三级生成！以上生成页数和占用空间是以全国省市县三级估算的，若是本地区分类请参考。</font>)
                  </td>
                </tr></td>
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
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
}
-->
</style>