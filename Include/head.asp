
<%Response.ContentType   =   "text/html"   
Response.CharSet   =   "gb2312"
Dim id,n,ci,c,c1,c2,c3,c4,dname,t,i,t1,t2,t3,n1,n2,n3,city_oneid,city_twoid,city_threeid,city_one,city_two,city_three,type_oneid,type_twoid,type_threeid,type_one,type_two,type_three,navigate,ic
'���ڶ�����������
'ȡ��HTTP�����ֵ����ֵ��HTOST��
Dim host,webhost
webhost=""
host=lcase(request.servervariables("HTTP_HOST"))
Call OpenConn

If cip=True then
'����IP�ж��Զ�������Ӧ����==========
Dim IPAdd,C_IP
C_IP=request.querystring("ip")
if trim(C_IP)="" then
   C_IP=Request.ServerVariables("REMOTE_ADDR")
elseif ubound(split(trim(ip),"."))<>3 then
   C_IP=Request.ServerVariables("REMOTE_ADDR") 'ip��ַ
end if
IPAdd=Look_Ip("ip/cz.dat",C_IP)

set rs=server.CreateObject("adodb.recordset")
rs.Open "select id,twoid,threeid,dname from DNJY_city where dname='"&host&"'",conn,1,1
if Not rs.EOF and not rs.BOF Then'aa
c1=trim(rs("id"))
c2=trim(rs("twoid"))
c3=trim(rs("threeid"))
c4=""&c1&","&c2&","&c3&""
Call CloseRs(rs)
else'bb
set rs=server.CreateObject("adodb.recordset")
rs.Open "select id,twoid,threeid,city,dname from DNJY_city ",conn,1,1
do while not rs.eof
if instr(IPAdd,rs("city"))>0 then
c1=trim(rs("id"))
c2=trim(rs("twoid"))
c3=trim(rs("threeid"))
c4=""&c1&","&c2&","&c3&""
Call CloseRs(rs)
exit Do
else'cc
c1=0
c2=0
c3=0
c4="0,0,0"
End If
rs.movenext
Loop
End If'dd
Call CloseRs(rs)
'����IP�ж��Զ�������Ӧ��������===========

else
set rs=server.CreateObject("adodb.recordset")
rs.Open "select id,twoid,threeid from DNJY_city where dname='"&host&"'",conn,1,1
if rs.EOF and rs.BOF Then
c1=0
c2=0
c3=0
c4="0,0,0"
Else
c1=trim(rs("id"))
c2=trim(rs("twoid"))
c3=trim(rs("threeid"))
c4=""&c1&","&c2&","&c3&""
End If
End If
Call CloseRs(rs)

c=trim(request("c"))
If trim(request("c"))="" Or trim(request("c"))="0" Then c=c4
c=split(c,",")
If c(0)="" Then c(0)=0'����ʾ���±�Խ�硱ʱ����config.asp�е�cip=True��Ϊfalse
c1=strint(c(0))
If ubound(c)<1 Then
c2=0
else
c2=strint(c(1))
End if
If ubound(c)<2 Then
c3=0
else
c3=strint(c(2))
End if


t=request("t")
If t="" Then t="0,0,0"
t=split(t,",")
If t(0)="" Then t(0)=0
t1=strint(t(0))
If ubound(t)<1 Then
t2=0
else
t2=strint(t(1))
End if
If ubound(t)<2 Then
t3=0
else
t3=strint(t(2))
End if

If c1="" Then c1=0
If c2="" Then c2=0
If c3="" Then c3=0
If t1="" Then t1=0
If t2="" Then t2=0
If t3="" Then t3=0
city_oneid=c1
city_twoid=c2
city_threeid=c3
type_oneid=t1
type_twoid=t2
type_threeid=t3

id=strint(request.QueryString("id"))

'

'������Ϣ��������жϣ����������� 
    Dim tttt
	IF t1=0 then
	tttt=""
	elseIF t3>0 then
	tttt=" and type_oneid="&t1&" and type_twoid="&t2&" and type_threeid="&t3
	elseIF t2>0 then
	tttt=" and type_oneid="&t1&" and type_twoid="&t2
	Else
	tttt=" and type_oneid="&t1
	End If
'========================================

'�����жϵ����л������������� 
    Dim ttcc,ttccnews,ttccdp,ttccmp,ttccbook,ttcc114,ttccbmxx,ttccads,ttcclink,ttcclogo 
    IF c3>0 Then
	ttcc=" and city_oneid="&c1&" and city_twoid="&c2&" and city_threeid="&c3
	ElseIF c2>0 Then
	ttcc=" and city_oneid="&c1&" and city_twoid="&c2
	ElseIF c1>0 Then
	ttcc=" and city_oneid="&c1
	End If

'���ݵ����жϾ�������鰴�����л���������������������������������
	ttccnews=ttcc
	ttccdp=ttcc
	ttccmp=ttcc
	ttccbook=ttcc
	ttcc114=ttcc
	ttccbmxx=ttcc
	ttccads=ttcc
	ttcclink=ttcc
	ttcclogo=ttcc
	If shows2="0" Then ttccnews=""
	If shows3="0" Then ttccdp=""
	If shows4="0" Then ttccmp=""
	If shows5="0" Then ttccbook=""
	If shows6="0" Then ttcc114=""
	If shows7="0" Then ttccbmxx=""
	If shows8="0" Then ttccads=""
	If shows9="0" Then ttcclink=""
	If shows10="0" Then ttcclogo=""
    If shows1="0" Then ttcc=""
	Session("ads")=ttccads'������
'================================================
'==================�����й���Ϣ==============================
Dim WebSetting,Tmpstr,ScriptName,WebTitle,WebText,belongtype
Dim cityweb,titleinfo,description,keywords,qq,qq_name,cityCompany,Cellular_phone,mymail,Tel,fax,serve,yzbm
Tmpstr = Split(Request.ServerVariables("PATH_INFO"),"/")
ScriptName = Lcase(Tmpstr(UBound(Tmpstr)))
If ScriptName="show.asp" or ScriptName="news_show.asp" or ScriptName="photo_show.asp" or ScriptName="General_stores.asp" Or ScriptName="com.asp" or ScriptName="company.asp" Then
set rs=server.createobject("adodb.recordset")
If ScriptName="show.asp" Then
sql = "select city_oneid,city_twoid,city_threeid,biaoti,keywords from DNJY_info where yz=1"
if jiaoyi=0 then
sql=sql&" and zt="&jiaoyi
end If
if overdue=0 Then
sql=sql&" and dqsj>"&DN_Now&""
end If
sql=sql&" and id="&cstr(id)
ElseIf ScriptName="General_stores.asp" Or ScriptName="com.asp" Then
m=CheckStr(trim(request("com")))
sql="select city_oneid,city_twoid,city_threeid,sdname,sdjj from [DNJY_user] where username='"&m&"'" 
ElseIf ScriptName="company.asp" Then
username=trim(request("username"))
sql="select city_oneid,city_twoid,city_threeid,mpname,mpjj from [DNJY_user] where mpkg=1 and username='"&username&"'"
ElseIf ScriptName="photo_show.asp" Then
owen=clng(request.querystring("id"))
sql="select city_oneid,city_twoid,city_threeid,title,keywords from [DNJY_photo] where show=1 and id="&owen&""&ttccnews&""
Else
owen=clng(request.querystring("id"))
sql="select city_oneid,city_twoid,city_threeid,title,keywords from DNJY_News where ifshow=1 and  id="&owen&""&ttccnews&""
	End If
rs.open sql,conn,1,1
	If Not Rs.Eof Then
		city_oneid=Rs(0)
		city_twoid=Rs(1)
		city_threeid=Rs(2)
		WebTitle=Rs(3)
		WebText=Rs(4)
	Else
	Call CloseRs(rs)
		Call CloseDB(conn)
		Response.Write("���������δ����˻������ڵ���")
		Response.End
	End If
	Call CloseRs(rs)
End If


	'�������ӵ���Ϣ���,���ڵ�ǰλ��
	IF type_oneid>0 Then
	set rs=server.createObject("ADODB.recordset")
	Rs.open "Select [name] from DNJY_type Where id="&type_oneid&" and twoid=0",conn
	IF not rs.eof then
	belongtype=Rs(0)
	IF type_twoid>0 Then
	rs.close
	Rs.open "Select [name] from DNJY_type Where id="&type_oneid&" and twoid="&type_twoid&" and threeid=0",conn
	IF not Rs.eof then
	belongtype=belongtype&" / "&Rs(0)
	IF type_threeid>0 Then
	rs.close
	Rs.open "Select [name] from DNJY_type Where id="&type_oneid&" and twoid="&type_twoid&" and threeid="&type_threeid&"",conn
	belongtype=belongtype&" / "&Rs(0)	
	End IF
	End IF	
	End IF
	End IF
	Call CloseRs(rs)
	End If
	
If  Not IsArray(Application(CacheName&"_WebSetting")) Then'Ĭ��Ϊ��վ
	Dim Temp
		Temp=Conn.Execute("Select Top 1 WebSetting,[help] From [DNJY_config]").GetRows
		Application.Lock			
			Application(CacheName&"_WebSetting")=Split(Temp(0,0),"|")			
			Application(CacheName&"_help")=Temp(1,0)
		    Application.unLock

End If
WebSetting=Application(CacheName&"_WebSetting")
cityweb=web
If city_oneid>0 Then'��վ����ʱʹ�÷�վ
	Sql="Select city,WebSetting,dname From [DNJY_city] Where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid
	Set Rs=Server.CreateObject("ADODB.recordset")
	Rs.open Sql,Conn,1,1
	If Rs.eof Then
	    Call CloseRs(rs)
		Call CloseDB(conn)
		Response.Write("û�е����������������޷�������ʾ��")
		Response.End
	Else	
		WebSetting = Rs(1)
		cityweb=Rs(2)
	End If
	Call CloseRs(rs)
If Len(WebSetting)<37 or Isnull(WebSetting) Then'���ֶ�Ϊ�ջ����зָ�����δ��������ʱ��ָ��վ��
WebSetting=""'����վδ�����κ�����ʱ��շ�վ����׼��ȡ��վ����
	For I=0 To 12'�������Ҫ����ĿΪֹ
		'If i<>4  And i<>6 Then WebSetting = WebSetting&Application(CacheName&"_WebSetting")(i) &"�ӡӡ�" '���˲���Ҫ��
		WebSetting = WebSetting&Application(CacheName&"_WebSetting")(i) &"�ӡӡ�" '����վ����Ϊ��ʱʹ����վ�ģ�������վ�������ݺ���ϡӡӡӣ��������վ����Ӧ����Ŀ��
	Next
	WebSetting =Left(WebSetting,Len(WebSetting)-3)'������������ӡӡ�
End If
WebSetting=Split(WebSetting,"�ӡӡ�")
End If
Dim ttkj_title,ttkj_keywords,ttkj_description
If Instr(Request.ServerVariables("URL"),"/news_show.asp") Or Instr(Request.ServerVariables("URL"),"/photo_show.asp") Or Instr(Request.ServerVariables("URL"),"/show.asp") Then
titleinfo=WebTitle
title=WebSetting(0)
description=titleinfo
keywords=WebText
ElseIf Instr(Request.ServerVariables("URL"),"/General_stores.asp") Or Instr(Request.ServerVariables("URL"),"/user/com.asp") Or Instr(Request.ServerVariables("URL"),"/company.asp") Then
titleinfo=WebTitle
title=WebSetting(0)
description=WebText
keywords=titleinfo
Else
title=WebSetting(0)
titleinfo=WebSetting(1)
description=WebSetting(2)
keywords=WebSetting(3)
End If

If Instr(Request.ServerVariables("URL"),"/xxindex.asp") Then titleinfo="������ϢƵ��"
If Instr(Request.ServerVariables("URL"),"/sdindex.asp") Then titleinfo="�̼ҵ���Ƶ��"
If Instr(Request.ServerVariables("URL"),"/qyindex.asp") Then titleinfo="��ҵ��ƬƵ��"
If Instr(Request.ServerVariables("URL"),"/news.asp") Or Instr(Request.ServerVariables("URL"),"/newssearch.asp") Then titleinfo="����/��Ѷ/չ��/Ƶ��"
If Instr(Request.ServerVariables("URL"),"/vipnews.asp") Or Instr(Request.ServerVariables("URL"),"/vipnewssearch.asp") Then titleinfo="���̹�����Դ"
If Instr(Request.ServerVariables("URL"),"/Message_board.asp") Then titleinfo="�û�����"
If Instr(Request.ServerVariables("URL"),"/help.asp") Then titleinfo="��������"


If Instr(Request.ServerVariables("URL"),"/list.asp") And request("xxsx")=1 Then belongtype="������Ϣ"
If Instr(Request.ServerVariables("URL"),"/list.asp") And request("xxsx")=2 Then belongtype="������Ϣ"
If Instr(Request.ServerVariables("URL"),"/list.asp") And request("xxsx")=3 Then belongtype="�Ƽ���Ϣ"
If Instr(Request.ServerVariables("URL"),"/list.asp") And request("xxsx")=4 Then belongtype="������Ϣ"

If Instr(Request.ServerVariables("URL"),"/news_show.asp") Or Instr(Request.ServerVariables("URL"),"/photo_show.asp") Or Instr(Request.ServerVariables("URL"),"/show.asp") Or Instr(Request.ServerVariables("URL"),"/General_stores.asp") Or Instr(Request.ServerVariables("URL"),"/company.asp") Then
ttkj_title=titleinfo&"_"&title
ttkj_keywords=keywords
ttkj_description=description
elseIf Instr(Request.ServerVariables("URL"),"/list.asp") Or Instr(Request.ServerVariables("URL"),"/list.asp") Then
ttkj_title=belongtype&"|"&title
ttkj_keywords=belongtype&"_"&keywords
ttkj_description=belongtype&"_"&title
Else
ttkj_title=title&"_"&titleinfo
ttkj_keywords=keywords
ttkj_description=description
End If


mymail=WebSetting(4)
Tel=WebSetting(5)
Cellular_phone=WebSetting(6)
fax=WebSetting(7)
serve=WebSetting(8)
yzbm=WebSetting(9)
qq=WebSetting(10)
qq_name=WebSetting(11)
cityCompany=WebSetting(12)


'==================�����й���ϢEND==============================

dim currentCity,belongcity
'���������л�ǰ
if city_oneid=0 then
	currentCity= "<b><font color=#ff000>��&nbsp;&nbsp;վ</font></b>"
	else
	currentCity= conn.execute ("select city from [DNJY_city] where id="&city_oneid&" and twoid=0")(0)
	if city_oneid>0 and city_twoid>0 then
	currentCity=conn.execute ("select city from [DNJY_city] where id="&city_oneid&" and twoid="&city_twoid&" and threeid=0")(0)
	end if
	if city_oneid>0 and city_twoid>0 and city_threeid>0 then
	currentCity=conn.execute ("select city from [DNJY_city] where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid&"")(0)
	end If
currentCity=CutStr(currentCity,22,"..")	
end if

	
'������������
navigate= " <A href=""http://"&web&"?c=0,0,0"">��վ</A> "
if city_oneid=0 then sql="select id,twoid,threeid,city,dname from DNJY_city where indexshow='yes' and id>0 and twoid=0 order by indexid asc"
if city_oneid>0 and city_twoid=0 then sql="select * from DNJY_city where indexshow='yes' and id="&city_oneid&" and twoid>0 and threeid=0 order by  indexid asc"
if city_oneid>0 and city_twoid>0 then sql="select * from DNJY_city where indexshow='yes' and id="&city_oneid&" and twoid="&city_twoid&" and threeid>0 order by  indexid asc"
Set rs=conn.execute(sql)
ic=1
do while not rs.eof
If rs("dname")<>"" And shows1=1 Then
navigate=navigate& " | <A href=""http://"&rs("dname")&""">"&rs("city")&"</A>"
dname=rs("dname")
else
navigate=navigate& " | <A href=""index.asp?c="&rs("id")&","&rs("twoid")&","&rs("threeid")&""">"&rs("city")&"</A>"
End if
if (ic mod citys)=0 then exit do
ic=ic+1
rs.movenext
loop
Call CloseRs(rs)
navigate=navigate& " | <A href=""city_list.asp"" target=""_blank"">����>>></A>"	


'���������л�
dim qiehuan,CityData
IF city_twoid>0 Then
Sql="select id,twoid,threeid,city,dname from DNJY_city where id="&city_oneid&" and twoid="&city_twoid&"  order by indexid asc"
ElseIF city_oneid>0 Then
Sql="select id,twoid,threeid,city,dname from DNJY_city where id="&city_oneid&" and threeid=0 order by indexid asc"
Else
Sql="select id,twoid,threeid,city,dname from DNJY_city where twoid=0 order by indexid asc"
End IF
Set Rs=conn.execute(Sql)
IF not Rs.eof then citydata=Rs.getrows
Call CloseRs(rs)
if isarray(citydata) then
qiehuan="<ul id=""h_qiehuan""><li><a href=""javascript:void(0)"" class=""ai""><IMG src=""img/city_more.gif"" width=60 height=21></a><br><div id=""qiehuanbox""><ul>"
for i=0 to ubound(citydata,2)
If citydata(4,i)<>"" Then
qiehuan=qiehuan&"<li><a href=""http://"&citydata(4,i)&""">"&citydata(3,i)&"</a></li>"
else
qiehuan=qiehuan&"<li><a href=""index.asp?c="&citydata(0,i)&","&citydata(1,i)&","&citydata(2,i)&""">"&citydata(3,i)&"</a></li>"
End if
next
qiehuan=qiehuan&"<li style=""float:right;width:300px;height:20px;""><a href=""http://"&web&"?c=0,0,0"">������վ</a> | <a href=""city_list.asp"">�������</a></li></ul></div></li></ul>"
end if	


Dim C_ID,T_ID,CT_ID,N_ID
C_ID="c="&c1&","&c2&","&c3
T_ID="t="&t1&","&t2&","&t3
CT_ID="c="&c1&","&c2&","&c3&"&t="&t1&","&t2&","&t3
N_ID="n="&n1&","&n2&","&n3
'��Ϣ���׼۸񣽣���������������������������������
function check_jiage(jiage)
if jiage=0 Then
check_jiage="����"
Else
check_jiage=""&jiage&"Ԫ"
end If
end Function
'��Ϣʱ����ʾ��ʽһ������������������������������������
function dicksj(sj_1) 
dim sj
sj=DateDiff("D",sj_1,now())
if sj<=1 then
response.write "<font color=""#FF0000"">����</font><img src=""images/new.gif"">"
elseif sj<=2 and sj>1 then
response.write "<font color=""#0000FF"">����"&timevalue(sj_1)&"</font>"
elseif sj>2 and sj<=3 then
response.write "<font color=""#008080"">ǰ��"&timevalue(sj_1)&"</font>"
elseif sj>3  then
response.write "<font color=#353535>"&datevalue(sj_1)&"</font>"
end if
end Function

'��Ϣʱ����ʾ��ʽ������������������������������
function dicksj2(sj_2) 
dim sj2,xxsj
sj2=DateDiff("D",sj_2,now())
if sj2<=1 then
dicksj2="<font color=""#FF0000"">����</font><img src=""images/new.gif"">"
elseif sj2<=2 and sj2>1 then
dicksj2= "<font color=""#0000FF"">����"&timevalue(sj_2)&"</font>"
elseif sj2>2 and sj2<=3 then
dicksj2="<font color=""#008080"">ǰ��"&timevalue(sj_2)&"</font>"
elseif sj2>3  then
dicksj2="<font color=#353535>"&datevalue(sj_2)&"</font>"
end If
end Function

 '�ж���Ϣ�Ƿ��о�̬�ļ����ڣ�����̬���ʣ�����������������������������������
Function x_path(p,id,pp,qx)'��վ����
If asphtm=1 And qx=0 then
	'Dim TempText,Fso'Ϊ���������ѹ��������ļ��Ƿ���ڣ�12������
	'TempText=Server.MapPath(p&"Html")
	'Set Fso = CreateObject("Scripting.FileSystemObject")
	'IF Fso.FileExists(TempText&"/"&ID&".html") Then
	'IF Fso.Getfile(TempText&"/"&ID&".html").Size>100 Then
	x_path=pp&"html/"&ID&".html"
	'Else
	'x_path=pp&"show.asp?id="&ID&""
	'End IF
	'Else
	'x_path=pp&"show.asp?id="&ID&""
	'End IF
	'Set Fso =Nothing
Else
x_path=pp&"show.asp?id="&ID&""
End if
End Function

 '�ж������Ƿ��о�̬�ļ����ڣ�����̬���ʣ�����������������������������������
Function S_path(p,id,pp,t)
    S_path=pp&"news_show.asp?Id="&ID'��δ���ɾ�̬���ܣ���ʱ�����
'���½�����˵
	'Dim TempText,Fso
	'TempText=Server.MapPath(p&"Html/news")
	'Set Fso = CreateObject("Scripting.FileSystemObject")
	'IF Fso.FileExists(TempText&"/"&ID&".html") Then
	'IF Fso.Getfile(TempText&"/"&ID&".html").Size>100 Then
	'S_path=pp&"html/news/"&ID&".html"
	'Else
	'S_path=pp&"news_show.asp?catid=0&bigcid=1&classid="&t&"&Id="&ID
	'End IF
	'Else
	'S_path=pp&"news_show.asp?catid=0&bigcid=1&classid="&t&"&Id="&ID
	'End IF
	'Set Fso =Nothing
'-------
End Function

 '�ж��Ƿ��з�վLOGO�ļ����ڣ�������վLOGO������������������������������������
Function l_path(p,logo)
	Dim TempText,Fso
	TempText=Server.MapPath(p)
	Set Fso = CreateObject("Scripting.FileSystemObject")
	IF Fso.FileExists(TempText&"/"&logo&".jpg") and shows10=1 Then	
	l_path=p&"/"&logo&".jpg"	
	Else
	l_path=p&"/0_0_0.jpg"
	End IF
	Set Fso =Nothing
End Function

 '�ж��Ƿ��վ��ʾ�����Ƿ��й��JS�ļ����ڣ�������ʾ��վ��棽����������������������������������
Function adjs_path(p,adid,jsid)
	Dim TempText,Fso
	TempText=Server.MapPath(p)
	Set Fso = CreateObject("Scripting.FileSystemObject")
	IF Fso.FileExists(TempText&"/"&adid&"_"&jsid&".js") and shows8=1 Then	
	adjs_path=p&"/"&adid&"_"&jsid&".js"		
	Else
	adjs_path=p&"/"&adid&"_0_0_0.js"	
	End IF
	Set Fso =Nothing
End Function

'�������������������ʼ�ȷ���ã�����������������������
'���EMAIL�Ƿ�ȷ���У�������ȷ������ַ���
Function IsqrEmail(email)
Dim ranid
	IsqrEmail=FALSE
	Set rs=Server.Createobject("adodb.recordset") 
	sql="select * from [DNJY_user] where email='"&email&"'"
	rs.open sql,conn,1,1
		If NOT(rs.bof AND rs.eof) Then 
			ranid=rs("ranid")
			IsqrEmail=TRUE
		End If
Call CloseRs(rs)
End Function

'�����漴�ַ���
Function makerndid(byVal maxLen)
	Dim strNewPass
	Dim whatsNext, upper, lower, intCounter
	RANdomize
	For intCounter = 1 To maxLen
	whatsNext = int(2 * Rnd)
	If whatsNext = 0 Then
		upper = 80 
		lower = 70
	Else
		upper = 48 
		lower = 39
	End If
	strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + upper))
	Next
	makerndid = strNewPass
End Function

'�����������ַ����Ƿ��ظ�
Function IsqrRanid(ranid)
	IsqrRanid=FALSE
	Set rs=Server.Createobject("adodb.recordset") 
	sql="select * from [DNJY_user] where ranid='"&Ranid&"'"
	rs.open sql,conn,1,1
		If NOT(rs.bof AND rs.eof) Then
			IsqrRanid=TRUE
		End If
Call CloseRs(rs)
End Function
'�������������������ʼ�ȷ����END������������������������

Call CloseDB(conn)%>
