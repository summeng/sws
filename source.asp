<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%dim server_v1,server_v2,errc
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))'����·��
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))'������·��
if mid(server_v1,8,len(server_v2))<>server_v2 then
response.write ("<table cellSpacing=0 borderColorDark=#ffffff cellPadding=0  width=600 align=center borderColorLight=#0066CC border=1 bgcolor='#F0F0F0'>")
response.write  ("<tr align=center>")
response.write    ("<td height=36>���ύ��·�����󣬽�ֹ��վ���ⲿ�ύ����,��رմ��ڣ�</td>")
response.write  ("</tr>")
response.write ("<tr align=center>")
response.write ("<form><td height=30>")
response.write ("<INPUT TYPE='BUTTON' value='�رմ���' onClick='window.close()'> ")
response.write ("</td></form>")
response.write ("</tr>")
response.write ("</table>")
response.end
end if
%>
<% 
if errc then 
response.write "<script language=""javascript"">" 
response.write "parent.alert('�Ƿ�����!��ֹ����վ�������!');" 
response.write "self.location.href='Index.Asp';" 
response.write "</script>" 
response.end 
end if 
%>
