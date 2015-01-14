<%
'====天天科技================================
' 关闭所有数据库对象并且释放内存 
On Error Resume Next 
if not isnull(conn) then
	conn.close
end if
if not isnull(rs) then
	rs.close
end if
'=============================================
%>