<!--#include file="../conn.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/Function.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("05")%>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
	margin-top: 16px;
}
-->
</style><body leftmargin="0">
<%Dim HtmlDir
HtmlDir = "/"&strInstallDir&"html/"

Response.Write "���ڸ�����Ϣ������״̬����"

Dim rsCreate, InfoPath, iCount, iTemp, TheFile, LastModifyTime, NeedUpdate,fso
iCount = strint(Conn.Execute("select count(0) from DNJY_info where yz=1 and fsohtm=1")(0))
Response.Write "һ����Ҫ��� " & iCount & " �����ݿ��б�ʶΪ�������ɡ�����Ϣ��"
iCount = strint(Conn.Execute("select count(0) from DNJY_info where  yz=1 and fsohtm=0")(0))
Response.Write "�������� " & iCount & " ����Ϣ��ʶΪ��δ���ɡ�������Ҫ�������״̬��<br>"

iCount = 0
iTemp = 0
Set Fso = Server.CreateObject("Scripting.FileSystemObject")
Set rsCreate = Conn.Execute("select id,fbsj from DNJY_info where yz=1 and fsohtm=1")
Do While Not rsCreate.EOF
    NeedUpdate = False
    InfoPath = ""&HtmlDir & rsCreate("id")&".html"
    If fso.FileExists(Server.MapPath(InfoPath)) Then
            Set TheFile = fso.GetFile(Server.MapPath(InfoPath))
            LastModifyTime = TheFile.DateLastModified
            If rsCreate("fbsj") > LastModifyTime Then
                NeedUpdate = True
            End If       
    Else
        NeedUpdate = True
    End If
    If NeedUpdate = True Then
        Conn.Execute ("update DNJY_info set fsohtm=0 where ID=" & rsCreate("ID") & "")
        iCount = iCount + 1
    End If
    iTemp = iTemp + 1
    If iTemp Mod 10 = 0 Then
        Response.Write "."
        Response.Flush
    End If
    If iTemp Mod 1000 = 0 Then
        Response.Write "<br>"
        Response.Flush
    End If
    rsCreate.MoveNext
Loop
rsCreate.Close
Set rsCreate = Nothing
Call CloseDB(conn)
Response.Write "<br><br>������Ϣ������״̬��ɣ�"
Response.Write "��鷢�ֹ��� " & iCount & "����Ϣʵ������δ���ɵģ��Ѿ�����������״̬��"
Response.Write "<p align='center'><a href='info_html.asp'>�����ء�</a></p>" & vbCrLf
%>
