<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%dim action
action=Request("action")
select case action
case "del"
if request.form("hfyz").count=0 then
response.write "<script language=javascript>alert('��û��ѡ��Ҫɾ���Ļظ���');history.go(-1);</script>"
response.End
end if
conn.execute ("delete from DNJY_reply where id in ("&request.form("hfyz")&")")
response.Write "<script language=javascript>alert('����ɾ���ɹ�!');location.replace('xxhflist.asp');</script>"
response.end
case "hfyz"
if request.form("hfyz").count=0 then
response.write "<script language=javascript>alert('��û��ѡ��Ҫ��˵Ļظ���');history.go(-1);</script>"
response.End
end if
conn.execute "update DNJY_reply set hfyz=1 where id in ("&request.form("hfyz")&")"
response.Write "<script language=javascript>alert('������˳ɹ�!');location.replace('xxhflist.asp');</script>"
response.end
case "delzhou"
dim theday
theday=date-7
conn.execute ("delete from DNJY_reply where hfsj<#"&theday&"# and hfyz=0")
response.Write "<script language=javascript>alert('һ��ǰδ��˻ظ�ɾ���ɹ�!');location.replace('xxhflist.asp');</script>"
response.end
case "delall"
conn.execute ("delete from DNJY_reply where hfyz=0")
response.Write "<script language=javascript>alert('����δ��˻ظ�ɾ���ɹ�!');location.replace('xxhflist.asp');</script>"
response.end
end select
%>