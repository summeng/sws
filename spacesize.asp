<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file=admin/cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("14")%>
<HTML>
<TITLE>系统空间占用</TITLE>
<LINK rel="stylesheet" href="admin/style.css" type="text/css">
<BODY topmargin="20">
<%
 	Sub ShowSpaceInfo(drvpath)
 		dim fso,d,size,showsize
 		set fso=server.createobject("scripting.filesystemobject")
 		drvpath=server.mappath(drvpath)
 		set d=fso.getfolder(drvpath)
 		size=d.size
 		showsize=size & "&nbsp;Byte"
 		if size>1024 then
 		   size=(size\1024)
 		   showsize=size & "&nbsp;KB"
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=formatnumber(size,2) & "&nbsp;MB"
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=formatnumber(size,2) & "&nbsp;GB"
 		end if
 		response.write "<font face=verdana>" & showsize & "</font>"
 	End Sub
 	Sub Showspecialspaceinfo(method)
 		dim fso,d,fc,f1,size,showsize,drvpath
 		set fso=server.createobject("scripting.filesystemobject")
 		drvpath=server.mappath("pic")
 		drvpath=left(drvpath,(instrrev(drvpath,"\")-1))
 		set d=fso.getfolder(drvpath)
 		if method="All" then
 			size=d.size
 		elseif method="Program" then
 			set fc=d.Files
 			for each f1 in fc
 				size=size+f1.size
 			next
 		end if
 		showsize=size & "&nbsp;Byte"
 		if size>1024 then
 		   size=(size\1024)
 		   showsize=size & "&nbsp;KB"
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=formatnumber(size,2) & "&nbsp;MB"
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=formatnumber(size,2) & "&nbsp;GB"
 		end if   
 		response.write "<font face=verdana>" & showsize & "</font>"
 	end sub
 	
 	Function Drawbar(drvpath)
 		dim fso,drvpathroot,d,size,totalsize,barsize
 		set fso=server.createobject("scripting.filesystemobject")
 		drvpathroot=server.mappath("pic")
 		drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
 		set d=fso.getfolder(drvpathroot)
 		totalsize=d.size
 		
 		drvpath=server.mappath(drvpath)
 		set d=fso.getfolder(drvpath)
 		size=d.size
 		
 		barsize=cint((size/totalsize)*400)
 		Drawbar=barsize
 	End Function
 	
 	Function Drawspecialbar()
 		dim fso,drvpathroot,d,fc,f1,size,totalsize,barsize
 		set fso=server.createobject("scripting.filesystemobject")
 		drvpathroot=server.mappath("pic")
 		drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
 		set d=fso.getfolder(drvpathroot)
 		totalsize=d.size
 		set fc=d.files
 		for each f1 in fc
 			size=size+f1.size
 		next
 		barsize=cint((size/totalsize)*400)
 		Drawspecialbar=barsize
 	End Function
 %>
 <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
    <TR>
    <TD height="20" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">系 统 空 间 的 使 用 情 况</FONT></TD>
  </TR>
  <tr>
  </table>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#ffffff">
<TR>
<TD width="447" height="140" valign=top >
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#ffffff">
<TR><TD> 		
<%Dim fsoflag,sPP,fso,f,fc,i,i2
fsoflag=1
if fsoflag=1 then%>
<BR>
<%
    Function GetPP
    dim s
    s=Request.ServerVariables("path_translated")
    GetPP=left(s,instrrev(s,"\",len(s)))
 			End function
 			if sPP="" then sPP=GetPP
 			if right(sPP,1)<>"\" then sPP=sPP&"\"
 			set fso=server.createobject("scripting.filesystemobject")
 			Set f = fso.GetFolder(sPP)
 			Set fc = f.SubFolders
 			i=1
                        i2=1
 			For Each f in fc
%>
目录<B><%=f.name%></B>占用空间：<IMG src="Images/Gaobei_VoteL.gif" width="8" height="9"><IMG src="Images/Gaobei_VoteM.gif" width="<%=drawbar(""&f.name&"")%>" height="9"><IMG src="Images/Gaobei_VoteR.gif" width="4" height="9">&nbsp;
<%showSpaceinfo(""&f.name&"")%><BR><BR>
<%
    i=i+1
    if i2<10 then
    i2=i2+1
    else
    i2=1
    end if
    Next
%>
 			程序文件占用空间：<IMG src="Images/Gaobei_VoteL.gif" width="8" height="9"><IMG src="Images/Gaobei_VoteM.gif" width="<%=drawspecialbar%>" height="9"><IMG src="Images/Gaobei_VoteR.gif" width="4" height="9">&nbsp;
<%showSpecialSpaceinfo("Program")%><BR><BR>
 			系统占用空间总计：<BR>
 			<IMG src="Images/Gaobei_VoteL.gif" width="8" height="9"><IMG src="Images/Gaobei_VoteM.gif" width="400" height="9"><IMG src="Images/Gaobei_VoteR.gif" width="4" height="9"> 
<%showspecialspaceinfo("All")%>
<%
    else
    response.write "<br><li>本功能已经被关闭"
    end if
%>
</TD>
</TR>
</TABLE></TD>
</TR>
</TABLE>
</BODY></HTML>