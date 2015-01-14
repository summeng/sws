<%
'=====以下参数必须保持正确，由后台设置生成，可用记事本手工修改，不明白就不要乱改动====www.ip126.com===

'==========================数据库类型====================================
Const SystemDatabaseType=0 '数据库类型 ACCESS数据库为0 SQL数据库为1 

'==========================Access数据库==================================
Const DBFileName="data/#ttgq.asp" 'ACCESS数据库存放的路径, 相对于根目录的路径

'==========================SQL数据库=====================================
Const SqlDatabaseName="ttgq" 'SQL数据库名(SqlDatabaseName)
Const SqlUsername="admin" 'SQL数据库登录用户名(SqlUsername)
Const SqlPassword="admin888"'SQL数据库用户登录密码(SqlPassword)
Const SqlHostIP="(local)"
'SQL主机地址(SqlHostIP)设置注意：程序与数据库同机用(local)；远程用IP，如"219.48.165.56"；数据库地址为域名的用"数据库域名"；本地调试时若与SQL Server其它版本并存安装时，SQL主机地址要这样填写："计算机名\实例名"，如"user\sql2008"
%>
