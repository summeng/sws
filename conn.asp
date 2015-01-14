<%@LANGUAGE="VBSCRIPT"%>
<%option explicit%>
<!--#include file="Database.asp"-->
<%
'==================================以下代码请勿改动==================================
dim startime,sys
startime=timer()
Dim Conn,rs,sql
Dim DN_True, DN_False, DN_Now, DN_OrderType, DN_DatePart_Y, DN_DatePart_M, DN_DatePart_D, DN_DatePart_H,DN_DatePart_N, DN_DatePart_S,DN_DatePart_W
Sub OpenConn() '定义打开数据库函数。在需要调用数据库的页面前加 Call OpenConn 调用执行以打开数据库，一般放在页面前面（数据库使用前）。  
    Dim ConnStr
    If SystemDatabaseType = 1 Then
        ConnStr = "Provider = Sqloledb; User ID = " & SqlUsername & "; Password = " & SqlPassword & "; Initial Catalog = " & SqlDatabaseName & "; Data Source = " & SqlHostIP & ";"
    Else
        ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/"&strInstallDir&""&DBFileName&"")		
    End If
    On Error Resume Next
    Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.open ConnStr
    If Err Then
        Err.Clear
        Set Conn = Nothing
        Response.Write "数据库连接出错，请检查Database.asp文件中的数据库参数设置。"
        Response.End
    End If
    If SystemDatabaseType = 1 Then
        DN_True = "1"
        DN_False = "0"
        DN_Now = "GetDate()"
        DN_OrderType = " desc"'倒序排序从大到小        
        DN_DatePart_Y = "yy"'年
        DN_DatePart_M = "mm"'月
		DN_DatePart_D = "dd"'日       
        DN_DatePart_H = "hh"'时
        DN_DatePart_N = "mi"'分
		DN_DatePart_S = "ss"'秒
        DN_DatePart_W = "wk"'一年第几周
    Else
        DN_True = "True"
        DN_False = "False"
        DN_Now = "Now()"
		DN_OrderType = " desc"'倒序排序从大到小             
        DN_DatePart_Y = "'yyyy'"'年
        DN_DatePart_M = "'m'"'月
		DN_DatePart_D = "'d'"'日       
        DN_DatePart_H = "'h'"'时
		DN_DatePart_N = "'n'"'分
		DN_DatePart_S = "'s'"'秒
		DN_DatePart_W = "'ww'"'一年第几周
    End If

End Sub

sub closers(crs)'定义关闭对象函数，等同 Rs.Close:Set Rs=Nothing。在操作结束时及时用 Call CloseRs(rs)
if Not crs Is Nothing Then '如果对象未清除 
if (crs.state and 1)=1 then
crs.close'释放对象
end if
set crs=nothing'清空内存
end if
end sub
sub CloseDB(Econn)'定义关闭数据库函数。等同conn.close:conn=nothing。在操作结束时及时用 Call CloseDB(conn) ，一般放在页面最后，节约资源。
if Not Econn is Nothing then'内存中存在
if (Econn.state and 1)=1 then
Econn.close'关闭连接
end if
set Econn=nothing'清空内存
end if
end Sub

sys=False%>