<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
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
str = str & "<description>"&Rankrs("city")&"������Ϣ</description>" & vbcrlf
set rs=server.createobject("ADODB.recordset")
sql = "select top 50 * from DNJY_info where yz=1 and fsohtm=1"&ttcc&""
if jiaoyi=0 Then sql=sql&" and zt="&jiaoyi
if overdue=0 Then sql=sql&" and dqsj>"&DN_Now&""
sql=sql&" order by fbsj "&DN_OrderType&""
rs.open sql,conn,1,1

if rs.Eof then
str = str & "<item>��δ��"&Rankrs("city")&"������Ϣ</item>" & vbcrlf
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
str = str & "<description><![CDATA[��ԱȨ���Ķ�]]></description>" & vbcrlf
elseIf rs("Readinfo")=2 Then
str = str & "<description><![CDATA[VIP��ԱȨ���Ķ�]]></description>" & vbcrlf
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
.SaveToFile server.mappath("/"&strInstallDir&"xml/rss"&Rankrs("id")&"_"&city_twoid&"_"&city_threeid&".xml"),2 '���ɵ�XML�ļ���,�������޸�,������и��õķ���,���Ը�
.Close
End With
Set objStream = Nothing
i=i+1
Rankrs.MoveNext          
loop  

If Not Err Then
response.Write "�ɹ���������RSS�ۺϷ���ҳ!������"&i&"ҳ"
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
    <TD height="20" bgcolor="#799AE1" align="center" colspan="13"><FONT color="#FFFFFF" style="font-size:14px">�� �� �� �� �� վ RSS �� ��</FONT></TD>
 </TR>
  <FORM name="form" method="post" action="?action=Production" >
  
    <tr> 
      <td height="30" colspan="3" align="center">&nbsp;</td>
    </tr>

    <tr> 
      <td width="188" height="25">               <tr>
                  <td width="20%" style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="33" bgcolor="#FFFFFF">
				  <p style="margin-left: 20">ѡ��������༶��<br>����һ��������������������
				
                  <td style="border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium" height="28" bgcolor="#FFFFFF"><p style="margin-left: 20">
                <input type="radio" name="Rank" value="1" id="Rank"  checked>
                 ȫ��һ�� (30��ҳ��&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="Rank" value="2" id="Rank">
                ȫ������ ��400��ҳ��&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="radio" name="Rank" value="3" id="Rank">
                ȫ������  ��2500��ҳ��<p>��<font color="#FF0000">��������ʱ��ռ�ô�����������Դ�������ɶ�ǧ��ҳ��ռ�ý�20M�ռ䣬�������������ɣ���������ҳ����ռ�ÿռ�����ȫ��ʡ������������ģ����Ǳ�����������ο���</font>)
                  </td>
                </tr></td>
    </tr>
    <tr> 
      <td height="30" colspan="3" align="center"> 
        <input type="submit" name="Submit" value="�ύ" class="input" >
        
        <input type="reset" name="Submit2" value="����" class="input"> 
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