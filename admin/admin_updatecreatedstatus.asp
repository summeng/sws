<!--#include file="../conn.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/Function.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
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

Response.Write "正在更新信息的生成状态……"

Dim rsCreate, InfoPath, iCount, iTemp, TheFile, LastModifyTime, NeedUpdate,fso
iCount = strint(Conn.Execute("select count(0) from DNJY_info where yz=1 and fsohtm=1")(0))
Response.Write "一共需要检查 " & iCount & " 条数据库中标识为“已生成”的信息，"
iCount = strint(Conn.Execute("select count(0) from DNJY_info where  yz=1 and fsohtm=0")(0))
Response.Write "其他还有 " & iCount & " 条信息标识为“未生成”，不需要检查生成状态。<br>"

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
Response.Write "<br><br>更新信息的生成状态完成！"
Response.Write "检查发现共有 " & iCount & "条信息实际上是未生成的，已经更新其生成状态。"
Response.Write "<p align='center'><a href='info_html.asp'>【返回】</a></p>" & vbCrLf
%>
