<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->

<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
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
	Response.Write "<li>文件（servercheck.asp）不存在！</li>"
    response.end
    End If
    if fso.fileExists(Server.MapPath("servercheck.asp")) then
fso.DeleteFile(Server.MapPath("servercheck.asp"))
end if
set fso=nothing
 
    If Err.Number <> 0 Then
        ErrMsg = ErrMsg & "<br><li>删除升级程序（update.asp）失败，错误原因：" & Err.Description & "<br>请手动删除此文件。"
        Err.Clear
    
    Else
        Response.Write "<li>删除（servercheck.asp）文件成功！</li>"
response.end
    End If
End If

response.write "为确保安全，不提供服务器环境探针，如需要测试服务器请与客服联系，切记用后要在admin目录下删除文件servercheck.asp！"
response.write "<p><a href='servercheck.asp'>请测试</a>&nbsp;&nbsp;&nbsp;&nbsp;<input name=""delfile"" type=""button"" value="" 立即删除servercheck.asp文件 "" onclick=""location='admin_server.asp?Action=del'"">"
response.end
dim founderr,errmsg
founderr=false
errmsg=""
if session("adminlogin")<>sessionvar then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你尚未登录，或者超时了！请<a href='index.asp'>重新登录</a>！"
  call diserror()
  response.end
end if
  if session("level")<>0 then
  errmsg=errmsg+"<br>"+"<li>你的管理权限不够！！</a>！"
  call diserror()
  response.end
  end if
%>
