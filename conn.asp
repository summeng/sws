<%@LANGUAGE="VBSCRIPT"%>
<%option explicit%>
<!--#include file="Database.asp"-->
<%
'==================================���´�������Ķ�==================================
dim startime,sys
startime=timer()
Dim Conn,rs,sql
Dim DN_True, DN_False, DN_Now, DN_OrderType, DN_DatePart_Y, DN_DatePart_M, DN_DatePart_D, DN_DatePart_H,DN_DatePart_N, DN_DatePart_S,DN_DatePart_W
Sub OpenConn() '��������ݿ⺯��������Ҫ�������ݿ��ҳ��ǰ�� Call OpenConn ����ִ���Դ����ݿ⣬һ�����ҳ��ǰ�棨���ݿ�ʹ��ǰ����  
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
        Response.Write "���ݿ����ӳ�������Database.asp�ļ��е����ݿ�������á�"
        Response.End
    End If
    If SystemDatabaseType = 1 Then
        DN_True = "1"
        DN_False = "0"
        DN_Now = "GetDate()"
        DN_OrderType = " desc"'��������Ӵ�С        
        DN_DatePart_Y = "yy"'��
        DN_DatePart_M = "mm"'��
		DN_DatePart_D = "dd"'��       
        DN_DatePart_H = "hh"'ʱ
        DN_DatePart_N = "mi"'��
		DN_DatePart_S = "ss"'��
        DN_DatePart_W = "wk"'һ��ڼ���
    Else
        DN_True = "True"
        DN_False = "False"
        DN_Now = "Now()"
		DN_OrderType = " desc"'��������Ӵ�С             
        DN_DatePart_Y = "'yyyy'"'��
        DN_DatePart_M = "'m'"'��
		DN_DatePart_D = "'d'"'��       
        DN_DatePart_H = "'h'"'ʱ
		DN_DatePart_N = "'n'"'��
		DN_DatePart_S = "'s'"'��
		DN_DatePart_W = "'ww'"'һ��ڼ���
    End If

End Sub

sub closers(crs)'����رն���������ͬ Rs.Close:Set Rs=Nothing���ڲ�������ʱ��ʱ�� Call CloseRs(rs)
if Not crs Is Nothing Then '�������δ��� 
if (crs.state and 1)=1 then
crs.close'�ͷŶ���
end if
set crs=nothing'����ڴ�
end if
end sub
sub CloseDB(Econn)'����ر����ݿ⺯������ͬconn.close:conn=nothing���ڲ�������ʱ��ʱ�� Call CloseDB(conn) ��һ�����ҳ����󣬽�Լ��Դ��
if Not Econn is Nothing then'�ڴ��д���
if (Econn.state and 1)=1 then
Econn.close'�ر�����
end if
set Econn=nothing'����ڴ�
end if
end Sub

sys=False%>