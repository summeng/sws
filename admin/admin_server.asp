<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->

<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("14")
dim Action
Action = Trim(request("Action"))
if Action="del" then
    On Error Resume Next
    Dim fso
    Set fso=Server.CreateObject("Scripting.FileSystemObject")
	if Not fso.fileExists(Server.MapPath("servercheck.asp")) Then
	Response.Write "<li>�ļ���servercheck.asp�������ڣ�</li>"
    response.end
    End If
    if fso.fileExists(Server.MapPath("servercheck.asp")) then
fso.DeleteFile(Server.MapPath("servercheck.asp"))
end if
set fso=nothing
 
    If Err.Number <> 0 Then
        ErrMsg = ErrMsg & "<br><li>ɾ����������update.asp��ʧ�ܣ�����ԭ��" & Err.Description & "<br>���ֶ�ɾ�����ļ���"
        Err.Clear
    
    Else
        Response.Write "<li>ɾ����servercheck.asp���ļ��ɹ���</li>"
response.end
    End If
End If

response.write "Ϊȷ����ȫ�����ṩ����������̽�룬����Ҫ���Է���������ͷ���ϵ���м��ú�Ҫ��adminĿ¼��ɾ���ļ�servercheck.asp��"
response.write "<p><a href='servercheck.asp'>�����</a>&nbsp;&nbsp;&nbsp;&nbsp;<input name=""delfile"" type=""button"" value="" ����ɾ��servercheck.asp�ļ� "" onclick=""location='admin_server.asp?Action=del'"">"
response.end
dim founderr,errmsg
founderr=false
errmsg=""
if session("adminlogin")<>sessionvar then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>����δ��¼�����߳�ʱ�ˣ���<a href='index.asp'>���µ�¼</a>��"
  call diserror()
  response.end
end if
  if session("level")<>0 then
  errmsg=errmsg+"<br>"+"<li>��Ĺ���Ȩ�޲�������</a>��"
  call diserror()
  response.end
  end if
%>
