<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=top.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>

<link href="/<%=strInstallDir%>css/css.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
</head>


<table width="980"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="21"><img src="img/top_s3.gif" width="21" height="49"></td>
        <td width="90" background="img/top_s1.gif"><div align="center"><img src="img/fangda.gif" width="70" height="27"></div></td>
        <td background="img/top_s1.gif"><%=f_search(3,3)%> </td>
        <td width="5"><img src="img/top_s2.gif" width="5" height="49"></td>
      </tr>
    </table></td>
  </tr>
</table>
<%
Dim searchname,theitem,thecity,thepage,currentpage,pagename,totalpage,page,itype,icity,cs,g,b,bb,endtime,ii,dqxx
city_twoid=0
city_threeid=0
type_twoid=0
type_threeid=0
Function   GetCount(S,C)   
          GetCount=0   
          Dim   I   
          Dim   Count  
		  count=0 
          For   I=1   to   Len(S)   
              If   Mid(S,I,1)=C   Then   Count=Count+1   
          Next   
          GetCount=Count   
      End   Function   

searchname=HtmlEncode(request("keyword"))
theitem=HtmlEncode(request("item"))
thecity=HtmlEncode(request("icity"))
itype=getcount(theitem,"a")
icity=getcount(thecity,"a")

leixing=HtmlEncode(request("leixing"))
If leixing="" Then
leixing=""
lx=""
Else
lx=leixing+"��"
leixing=" And leixing='"&leixing&"'"
End if

Set rs = Server.CreateObject("ADODB.Recordset")
  'if searchname="" or searchname="�����������ؼ���" then'��Ҫ��ؼ���
    'response.write "<SCRIPT language=JavaScript>alert('�����������ؼ��֣�');"
	'response.write "this.location.href='javascript:history.back()';</SCRIPT>"
   	'response.end
  'Else
  select case itype'��Ϣ����
  case "0"
  type_oneid=theitem
  case "1"
  i=InStr(theitem,"a")
  type_oneid=left(theitem, i-1)
  ii=len(theitem)
  type_twoid=Right(theitem,ii-i)
  case "2"
  i=InStr(theitem,"a")
  type_oneid=left(theitem, i-1)
  ii=len(theitem)
  theitem=Right(theitem,ii-i)
  i=InStr(theitem,"a")
  type_twoid=left(theitem, i-1)
  ii=len(theitem)
  type_threeid=Right(theitem,ii-i)
  end select
'���������������������б�ʼ,Ҫ��label.asp��������ϣ���������
'------------���������б�,��������ʾ��ʼ���벻����������ͬʱ��------------
  select case icity'��������
  case "0"
  city_oneid=thecity
  case "1"
  i=InStr(thecity,"a")
  city_oneid=left(thecity, i-1)
  ii=len(thecity)
  city_twoid=Right(thecity,ii-i)
  case "2"
  i=InStr(thecity,"a")
  city_oneid=left(thecity, i-1)
  ii=len(thecity)
  thecity=Right(thecity,ii-i)
  i=InStr(thecity,"a")
  city_twoid=left(thecity, i-1)
  ii=len(thecity)
  city_threeid=Right(thecity,ii-i)
  end Select
'------------���������б�,��������ʾ��ʼ���벻����������ͬʱ��END------------
'------------���������б�,������������ʾȫ����ʼ-----------------------
'If HtmlEncode(request("icity"))<>"��վ" Then
'cs=split(thecity,"a")
'response.write cs(1)
'response.end
'city_oneid=strint(cs(0))
'city_twoid=strint(cs(1))
'city_threeid=strint(cs(2))
'End if
'------------���������б�,������������ʾȫ����ʼEND-----------------------
'���������������������б�ʼ,Ҫ��label.asp���������END����������
If overdue=0 Then'���ݹ�����Ϣ�Ƿ���ʾ������
dqxx=" and dqsj >= "&DN_Now&""
Else
dqxx=""
End if
    if theitem="ȫ��" and thecity="��վ" then
    sql="select * from DNJY_info where (biaoti Like '%"& searchname &"%' or memo Like '%"& searchname &"%') and yz=1"&dqxx&""&leixing&" order by hfje/fbts "&DN_OrderType&",b "&DN_OrderType&",id "&DN_OrderType&""
    rs.open sql,conn,1,1
    end if
	if theitem<>"ȫ��" and thecity="��վ" then
	if type_oneid<>0 and type_twoid=0 and type_threeid=0 then 
    sql="select * from DNJY_info where (biaoti Like '%"& searchname &"%' or memo Like '%"& searchname &"%') and type_oneid="&type_oneid&" and yz=1"&dqxx&""&leixing&" order by hfje/fbts "&DN_OrderType&",b "&DN_OrderType&",id "&DN_OrderType&""
	end if
	if type_oneid<>0 and type_twoid<>0 and type_threeid=0 then 
	sql="select * from DNJY_info where (biaoti Like '%"& searchname &"%' or memo Like '%"& searchname &"%') and type_oneid="&type_oneid&" and type_twoid="&type_twoid&" and yz=1"&dqxx&""&leixing&" order by hfje/fbts "&DN_OrderType&",b "&DN_OrderType&",id "&DN_OrderType&""
	end if
	if type_oneid<>0 and type_twoid<>0 and type_threeid<>0 then 
	sql="select * from DNJY_info where (biaoti Like '%"& searchname &"%' or memo Like '%"& searchname &"%') and type_oneid="&type_oneid&" and type_twoid="&type_twoid&" and type_threeid="&type_threeid&"and yz=1"&dqxx&""&leixing&" order by hfje/fbts "&DN_OrderType&",b "&DN_OrderType&",id "&DN_OrderType&""
	end if
	rs.open sql,conn,1,1
	end if
	if theitem="ȫ��" and thecity<>"��վ" then
	if city_oneid<>0 and city_twoid=0 and city_threeid=0 then 
    sql="select * from DNJY_info where (biaoti Like '%"& searchname &"%' or memo Like '%"& searchname &"%') and city_oneid="&city_oneid&" and yz=1"&dqxx&""&leixing&" order by hfje/fbts "&DN_OrderType&",b "&DN_OrderType&",id "&DN_OrderType&""
	end if
	if city_oneid<>0 and city_twoid<>0 and city_threeid=0 then 
	sql="select * from DNJY_info where (biaoti Like '%"& searchname &"%' or memo Like '%"& searchname &"%') and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and yz=1"&dqxx&""&leixing&" order by hfje/fbts "&DN_OrderType&",b "&DN_OrderType&",id "&DN_OrderType&""
	end if
	if city_oneid<>0 and city_twoid<>0 and city_threeid<>0 then 
	sql="select * from DNJY_info where (biaoti Like '%"& searchname &"%' or memo Like '%"& searchname &"%') and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid&" and yz=1"&dqxx&""&leixing&" order by hfje/fbts "&DN_OrderType&",b "&DN_OrderType&",id "&DN_OrderType&""
	end if
	rs.open sql,conn,1,1
    end if
	if theitem<>"ȫ��" and thecity<>"��վ" then
	if type_oneid<>0 and type_twoid=0 and type_threeid=0 then
	if city_oneid<>0 and city_twoid=0 and city_threeid=0 then
	sql="select * from DNJY_info where (biaoti Like '%"& searchname &"%' or memo Like '%"& searchname &"%') and type_oneid="&type_oneid&" and city_oneid="&city_oneid&" and yz=1"&dqxx&""&leixing&" order by hfje/fbts "&DN_OrderType&",b "&DN_OrderType&",id "&DN_OrderType&""
    end if
	if city_oneid<>0 and city_twoid<>0 and city_threeid=0 then
	sql="select * from DNJY_info where (biaoti Like '%"& searchname &"%' or memo Like '%"& searchname &"%') and type_oneid="&type_oneid&" and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and yz=1"&dqxx&""&leixing&" order by hfje/fbts "&DN_OrderType&",b "&DN_OrderType&",id "&DN_OrderType&""
    end if
	if city_oneid<>0 and city_twoid<>0 and city_threeid<>0 then
	sql="select * from DNJY_info where (biaoti Like '%"& searchname &"%' or memo Like '%"& searchname &"%') and type_oneid="&type_oneid&" and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid&" and yz=1"&dqxx&""&leixing&" order by hfje/fbts "&DN_OrderType&",b "&DN_OrderType&",id "&DN_OrderType&""
    end if
	end if
	
	if type_oneid<>0 and type_twoid<>0 and type_threeid=0 then
	if city_oneid<>0 and city_twoid=0 and city_threeid=0 then
	sql="select * from DNJY_info where (biaoti Like '%"& searchname &"%' or memo Like '%"& searchname &"%') and type_oneid="&type_oneid&" and type_twoid="&type_twoid&" and city_oneid="&city_oneid&" and yz=1"&dqxx&""&leixing&" order by hfje/fbts "&DN_OrderType&",b "&DN_OrderType&",id "&DN_OrderType&""
	end if
	if city_oneid<>0 and city_twoid<>0 and city_threeid=0 then
	sql="select * from DNJY_info where (biaoti Like '%"& searchname &"%' or memo Like '%"& searchname &"%') and type_oneid="&type_oneid&" and type_twoid="&type_twoid&" and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and yz=1"&dqxx&""&leixing&" order by hfje/fbts "&DN_OrderType&",b "&DN_OrderType&",id "&DN_OrderType&""
	end if
	if city_oneid<>0 and city_twoid<>0 and city_threeid<>0 then
	sql="select * from DNJY_info where (biaoti Like '%"& searchname &"%' or memo Like '%"& searchname &"%') and type_oneid="&type_oneid&" and type_twoid="&type_twoid&" and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid&" and yz=1"&dqxx&""&leixing&" order by hfje/fbts "&DN_OrderType&",b "&DN_OrderType&",id "&DN_OrderType&""
	end if
	end if
	
	if type_oneid<>0 and type_twoid<>0 and type_threeid<>0 then
	if city_oneid<>0 and city_twoid=0 and city_threeid=0 then
	sql="select * from DNJY_info where (biaoti Like '%"& searchname &"%' or memo Like '%"& searchname &"%') and type_oneid="&type_oneid&" and type_twoid="&type_twoid&" and type_threeid="&type_threeid&" and city_oneid="&city_oneid&" and yz=1"&dqxx&""&leixing&" order by hfje/fbts "&DN_OrderType&",b "&DN_OrderType&",id "&DN_OrderType&""
	end if
	if city_oneid<>0 and city_twoid<>0 and city_threeid=0 then
	sql="select * from DNJY_info where (biaoti Like '%"& searchname &"%' or memo Like '%"& searchname &"%') and type_oneid="&type_oneid&" and type_twoid="&type_twoid&" and type_threeid="&type_threeid&" and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and yz=1"&dqxx&""&leixing&" order by hfje/fbts "&DN_OrderType&",b "&DN_OrderType&",id "&DN_OrderType&""
	end if
	if city_oneid<>0 and city_twoid<>0 and city_threeid<>0 then
	sql="select * from DNJY_info where (biaoti Like '%"& searchname &"%' or memo Like '%"& searchname &"%') and type_oneid="&type_oneid&" and type_twoid="&type_twoid&" and type_threeid="&type_threeid&" and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid&" and yz=1"&dqxx&""&leixing&" order by hfje/fbts "&DN_OrderType&",b "&DN_OrderType&",id "&DN_OrderType&""
	end if
	end if
	rs.open sql,conn,1,1
    end if
'end if'��Ҫ��ؼ���

if rs.eof then
	response.write "<SCRIPT language=JavaScript>alert('�Բ���û���ҵ�������������Ϣ��');"
	response.write "this.location.href='javascript:history.back()';</SCRIPT>"
	rs.close
	set rs=nothing
   	response.end
end if
thepage=request("pageid")
if isnumeric(thepage)=false then
response.write "<script>alert('�������󣬴��ڹرգ�');history.back();</script>"
response.end
end if
currentpage = cint(thepage)            
if currentpage <= 0 then currentpage=1       
const maxperpage=10
rs.Pagesize = maxperpage  
pagename="search.asp"                     
totalpage = rs.PageCount 
rs.absolutepage = currentpage 
n=rs.recordcount                    
page = PageSet(currentpage,totalpage,pagename)
endtime=timer()%>
<TABLE WIDTH=980 BORDER=0 CELLPADDING=0 CELLSPACING=0 align="center" bgcolor="fff000">
           <TR> 
             <TD width="100%" height="30">
               <P align="center"><FONT style="font-size:14px"><B>������������ؼ��֣�<FONT color="red"><%=searchname%></FONT>&nbsp;&nbsp;���������������</B></FONT></P>                
             </TD>                                          
           </TR>
</TABLE>


<TABLE width="980" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="5ABE06">
  <TR>
    <TD height="26" align="center" bgcolor="E7FFB5"> ��������<B><%=n%></B>����<FONT color="red"><b><%=searchname%></b></FONT>��ص�<b><%=lx%></b>��Ϣ&nbsp;&nbsp;<%response.write "������ʱ��<b>"&FormatNumber((endtime-startime),3)&"</b>��"%>&nbsp;&nbsp;ÿҳ<B><%=maxperpage%></B>��&nbsp;&nbsp;ҳ�Σ�<B><%=currentpage%>/<%=totalpage%></B>&nbsp;&nbsp;&nbsp;&nbsp;
    <%response.write page%></TD>
  </TR>
</TABLE>
<TABLE  border="0" cellpadding="0" cellspacing="0" align="center" id="table23">
  <TR>
    <TD width="10" height="6"></TD>
  </TR>
</TABLE>

<TABLE width="980" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="ffffff">
<TR><TD width="10"></TD>
	<TD width="620">
<!--�����뿪ʼ-->
  <script src="<%=adjs_path("ads/js","468",c1&"_"&c2&"_"&c3)%>"></script>
  <!--���������-->
<%i=0 
do while i < maxperpage and not rs.eof%><HR>
	<table border="0" cellpadding="0" cellspacing="0" ><tr><td class=f><%if rs("Readinfo")=1 Then Response.Write "<img src=""img/lock.jpg"" alt=""��ͨ��ԱȨ�����"">"%><%if rs("Readinfo")=2 Then Response.Write "<img src=""img/lockvip.jpg"" alt=""VIP��ԱȨ�����"">"%><a href="<%=x_path("",rs("id"),"",rs("Readinfo"))%>" target="_blank"><font size="3"><b><% 
response.write Replace (rs("biaoti"),searchname,"<font color=#FF0000>" & searchname & "</font>") 
%> </a></b><%if rs("tupian")<>"0" Then response.write "<img src=""images/num/pic.gif"" alt=""��ͼƬ"">"
if rs("tj")>0 Then response.write"<img src=""images/jian.gif"" alt=""�Ƽ���Ϣ"" width=""15"" height=""15"">" 
b=rs("b")
bb=len(b)
if b<>0 Then 
response.write "<img src=""images/num/jsq.gif"" alt=""�ö�"&b&"��"">"
for g=1 to bb
response.write"<img src=""images/num/"&Mid(b,g,1)&".gif"" alt=""�ö�"&b&"��"">"
next
end if
%>
<br><font size=-1><% 
Dim vip,memo,dianhua
memo=CutStr(RemoveHTML_ttkj(trim(rs("memo"))),600,"...")
dianhua=rs("dianhua")
If request.cookies("dick")("username")<>rs("username") Then'�Լ�������ȫȨ�Ķ�
vip=0'��¼��Ա
If request.cookies("dick")("username")<>"" Then vip=conn.Execute("Select vip From DNJY_user Where username='"&request.cookies("dick")("username")&"'")(0)
If rs("Readinfo")=1 And request.cookies("dick")("username")="" Then
memo="<font color=#666666>��ԱȨ���Ķ�</font>"
dianhua="<font color=#666666>��ԱȨ���Ķ�</font>"
End If
If rs("Readinfo")=2 And vip<1 Then
memo="<font color=#666666>VIP��ԱȨ���Ķ�</font>"
dianhua="<font color=#666666>VIP��ԱȨ���Ķ�</font>"
End If
End If
response.write replace(memo,searchname,"<font color=#FF0000>" & searchname & "</font>") 
%> 
 <br><font color=#008000>(��ϵ��:<%=rs("username")%>|�绰:<%=dianhua%>|��Ч��:<%=datediff("d",now(),rs("dqsj"))%>��)</font>  <a href="<%=x_path("",rs("id"),"",rs("Readinfo"))%>" target="_blank"><u>��վ����</u></a><br>
 <%	belongcity="<a href=""http://"&web&"/"&strInstallDir&"index.asp?c="&rs("city_oneid")&",0,0"">"&rs("city_one")&"</a>"
    IF rs("city_two")<>"" and not isnull(rs("city_two")) Then belongcity=belongcity&"/<a href=""http://"&web&"/"&strInstallDir&"index.asp?c="&rs("city_oneid")&","&rs("city_twoid")&",0"">"&rs("city_two")&"</a>"
	IF rs("city_three")<>"" and not isnull(rs("city_three")) Then belongcity=belongcity&"/<a href=""http://"&web&"/"&strInstallDir&"index.asp?c="&rs("city_oneid")&","&rs("city_twoid")&","&rs("city_threeid")&""">"&rs("city_three")&"</a>"
	response.write belongcity
%> | <%
	belongtype="<a href=""http://"&web&"/"&strInstallDir&"list.asp?t="&rs("type_oneid")&",0,0&"&C_ID&""">"&rs("type_one")&"</a>"
	IF rs("type_two")<>"" and not isnull(rs("type_two")) Then belongtype=belongtype&"/<a href=""http://"&web&"/"&strInstallDir&"list.asp?t="&rs("type_oneid")&","&rs("type_twoid")&",0&"&C_ID&""">"&rs("type_two")&"</a>"
	IF rs("type_three")<>"" and not isnull(rs("type_three")) Then belongtype=belongtype&"/<a href=""http://"&web&"/"&strInstallDir&"list.asp?t="&rs("type_oneid")&","&rs("type_twoid")&","&rs("type_threeid")&"&"&C_ID&""">"&rs("type_three")&"</a>"
	response.write belongtype%></font></td></tr></table><br>
<%i=i+1
if i>=rs.Pagesize then exit do
rs.movenext
Loop
rs.close:set rs=nothing                                                 
%>  

	</TD>
	<TD width="23"></TD>
	<TD width="300" valign="top"><div style="border-left:1px solid #e1e1e1;padding-left:10px;word-break:break-all;word-wrap:break-word;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>������Ϣ</b><p><%=f_info(2,0,0,0,c1,c2,c3,0,0,0,0,0,1,1,0,0,0,26,20,1,13,11)%></div></TD>
</TR>
</TABLE>
<%Dim  strSplit,Sp,Pagestart,Pageend,j                                                                 
Function PageSet(currentpage,totalpage,pagename) 
if currentpage mod 10 = 0 then    
Sp = currentpage \ 10      
else                                                                                                  
Sp = currentpage \ 10 + 1 
end if                                                                                                  
Pagestart = (Sp-1)*10+1                                                                                                  
Pageend = Sp*10                                                                                                  
strSplit = "<a href="&pagename&"?keyword="&searchname&"&item="&theitem&"&icity="&thecity&"&pageid=1><font face=webdings title=��һҳ>9</font></a>&nbsp;"                                                                                                  
if Sp > 1 then strSplit = strSplit & "<a href="&pagename&"?keyword="&searchname&"&item="&theitem&"&icity="&thecity&"&pageid="&Pagestart-10&"><font face=webdings title=ǰʮҳ>7</font></a>&nbsp;"                                                                                                  
for j=PageStart to Pageend                                                                                                  
if j > totalpage then exit for                                                                                                  
if j <> currentpage then                                                                                                  
strSplit = strSplit & "<a href="&pagename&"?keyword="&searchname&"&item="&theitem&"&icity="&thecity&"&pageid="&j&">["&j&"]</a>&nbsp;"			                                                                                                  
else                                                                                                  
strSplit = strSplit & "<font color=red>["&j&"]</font>&nbsp;"                                                                                                  
end if                                                                                                  
next                                                                                                  
if Sp*10  < totalpage then strSplit = strSplit & "<a href="&pagename&"?keyword="&searchname&"&item="&theitem&"&icity="&thecity&"&pageid="&Pagestart+10&"><font face=webdings title=��ʮҳ>8</font></a>"                                                                                                   
strSplit = strSplit & "<a href="&pagename&"?keyword="&searchname&"&item="&theitem&"&icity="&thecity&"&pageid="&totalpage&" ><font face=webdings title=""���һҳ"">:</font></a>"                                                                                                  
PageSet = strSplit                                                                                              
End Function                                                                       
%>  

<TABLE  border="0" cellpadding="0" cellspacing="0" align="center" id="table23">
  <TR>
    <TD height="6"></TD>
  </TR>
</TABLE>
<TABLE width="980" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="5ABE06">
  <TR>
    <TD height="26" align="center" bgcolor="E7FFB5">��������<B><%=n%></B>����<FONT color="red"><b><%=searchname%></b></FONT>��ص�<b><%=lx%></b>��Ϣ&nbsp;&nbsp;<%response.write "������ʱ��<b>"&FormatNumber((endtime-startime),3)&"</b>��"%>&nbsp;&nbsp;&nbsp;ÿҳ<B><%=maxperpage%></B>��&nbsp;&nbsp;&nbsp;ҳ�Σ�<B><%=currentpage%>/<%=totalpage%></B>&nbsp;&nbsp;&nbsp;
    <%response.write page%></TD>
  </TR>
</TABLE>
<TABLE border="0" width="980" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF" align=center>
  <TR>
  </TR>
</TABLE>

<!--#include file="footer.asp"-->

