<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%dim action
action=Request("action")
select case action
case "del"
if request.form("hfyz").count=0 then
response.write "<script language=javascript>alert('您没有选择要删除的回复？');history.go(-1);</script>"
response.End
end if
conn.execute ("delete from DNJY_reply where id in ("&request.form("hfyz")&")")
response.Write "<script language=javascript>alert('批量删除成功!');location.replace('xxhflist.asp');</script>"
response.end
case "hfyz"
if request.form("hfyz").count=0 then
response.write "<script language=javascript>alert('您没有选择要审核的回复？');history.go(-1);</script>"
response.End
end if
conn.execute "update DNJY_reply set hfyz=1 where id in ("&request.form("hfyz")&")"
response.Write "<script language=javascript>alert('批量审核成功!');location.replace('xxhflist.asp');</script>"
response.end
case "delzhou"
dim theday
theday=date-7
conn.execute ("delete from DNJY_reply where hfsj<#"&theday&"# and hfyz=0")
response.Write "<script language=javascript>alert('一周前未审核回复删除成功!');location.replace('xxhflist.asp');</script>"
response.end
case "delall"
conn.execute ("delete from DNJY_reply where hfyz=0")
response.Write "<script language=javascript>alert('所有未审核回复删除成功!');location.replace('xxhflist.asp');</script>"
response.end
end select
%>