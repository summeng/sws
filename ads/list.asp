<!--#include file="../conn.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("13")%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>COAD System</title>
<SCRIPT LANGUAGE="JavaScript">
function delad(){
if (confirm("ɾ����潫ͬʱɾ�����ͼƬ��JS���룬�����Իָ���ȷ��ɾ����")){return true;}
return false;
}
</SCRIPT>
<script language=javascript src=../Include/mouse_on_title.js></script>
<link href="../css/style.css" rel="stylesheet" type="text/css">

<style type="text/css">
<!--
.style1 {color: #FFFFFF;
	font-weight: bold;
}
-->
</style>
<!---������忪ʼ--->
<script>
function loadThreadFollow(t_id,b_id){
	var targetImg =eval("document.all.followImg" + t_id);
	var targetDiv =eval("document.all.follow" + t_id);
	
	if ("object"==typeof(targetImg)){
		if (targetDiv.style.display!='block'){
			targetDiv.style.display="block";
			
		}else{
			targetDiv.style.display="none";
		}
	}
}
</script>
<!---����������--->
</head>
<body>
<%Call OpenConn
Dim oneid,twoid,threeid,Rsgg,Data,dbpath,rsc,city,ADhost,ADhost1,ADhost2,ADhost3,ADhost4,ADhost5,ADhost6,ADhost7,ADhost8,ADhost9
oneid=strint(Request("oneid"))
twoid=strint(Request("twoid"))
threeid=strint(Request("threeid"))
If Request("oneid")="" Then oneid=0
If Request("twoid")="" Then twoid=0
If Request("threeid")="" Then threeid=0
city=" and city_oneid="&oneid&" and city_twoid="&twoid&" and city_threeid="&threeid&""
if Request("all")="ok" Then city=""
Temp="oneid="&oneid&"&twoid="&twoid&"&threeid="&threeid
Dim pagesize,curpage,strcate,skin
Dim ADID,ADViews,ADHits,ADType,ADSrc,ADLink,ADAlt,ADWidth,ADHeight,ADNote,ADStopViews,ADStopHits,ADStopDate,IsSys,adkg
Set rs = Server.CreateObject("ADODB.Recordset")
if Request.QueryString("action")="stop" then
	sql = "SELECT * FROM [DNJY_ad] where ( ADStopViews <> 0 and ADViews > ADStopViews"&city&") or ( ADStopHits <> 0 and ADHits > ADStopHits"&city&") or ( DateDiff("&DN_DatePart_D&","&DN_Now&",ADStopDate)<1"&city&" ) ORDER BY id "&DN_OrderType&""
Else
	sql = "SELECT * FROM [DNJY_ad] where ADID<>''"&city&" ORDER BY id "&DN_OrderType&""
end If
rs.open sql, conn, 1, 1

'����Ƿ���ڴ���
If err.number <> 0 Then
	Response.Write "�����ݿ����,����conn.asp�����ݿ��ַ�Ƿ���ȷ"
	Response.End
End If

If rs.bof or rs.eof Then
	rs.Close
	if Request.QueryString("action")="stop" Then
	%>
	<table width="100%" height="200" border="0" align="center" cellpadding="0" cellspacing="1">
  <tr>
    <td align="center"><font color="#990000">��ʱû�й��ڹ�棡</font></td>
  </tr>
</table>
	<%Else%>
	<table width="100%" height="200" border="0" align="center" cellpadding="0" cellspacing="1">
  <tr>
    <td align="center" bgcolor="#DEDBEF">��ʱû�з���������棡-- <a href=add.asp>����<font color="#0000ff">����¹��</font></a></td>
  </tr>
</table>
	<%
	End If
	Response.End
End If

%>
<%'=============��ҳ���忪ʼ���ɷ������ݿ��ǰ���
                 dim action
    action=request.QueryString("action")   
    Const MaxPerPage=10   '����ÿҳ��ʾ��¼�����ɸ���ʵ���Զ���
       dim totalPut   
       dim CurrentPage
       dim TotalPages
       dim j

        if Not isempty(request("page")) And request("page")<>"" then
          currentPage=Cint(request("page"))
       else
          currentPage=1
       end if        
'=============��ҳ�������%>
<%'=============��ҳ����뿪ʼ����������ݿ����ݱ�򿪺�
   
    if err.number<>0 then
    response.write "<p align='center'>���ݿ�����ʱ������</p>"
    end if    
      if rs.eof And rs.bof then
           Response.Write "<p align='center'>�Բ���û�з���������¼��</p>"
       else
      totalPut=rs.recordcount
          if currentpage<1 then
              currentpage=1
          end if

          if (currentpage-1)*MaxPerPage>totalput then
         if (totalPut mod MaxPerPage)=0 then
           currentpage= totalPut \ MaxPerPage
         else
            currentpage= totalPut \ MaxPerPage + 1
         end if
          end if

           if currentPage=1 then
               showContent               
               
               showpage totalput,MaxPerPage,""&request.ServerVariables("script_name")&""  
           else
              if (currentPage-1)*MaxPerPage<totalPut then
                rs.move  (currentPage-1)*MaxPerPage
                dim bookmark
                bookmark=rs.bookmark
                showContent
                 showpage totalput,MaxPerPage,""&request.ServerVariables("script_name")&""  
            else
             currentPage=1
                showContent
                
                showpage totalput,MaxPerPage,""&request.ServerVariables("script_name")&""  
                
           end if
        end if
           end if
'=============��ҳ��������%>

 <%'=============ѭ���忪ʼ
   sub showContent
   dim i
   i=0  %>
<table width='100%' height="22" border=0 align=center cellpadding=0 cellspacing=1 bgcolor="#6699cc">
<tr bgcolor="#6699cc">
  <td height="28" align="center"><div align="left">
    <table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr><td height="28" colspan="2"><div align="center" ><font color="#FFFFFF"><b>����б�</b></font></div></td>
        
        <td>
              <div align="right">
               
              </div></td>
      </tr>
    </table>
  </div>        </td>
  </tr>
</table>          <TR>
          <TD width="60%">&nbsp;&nbsp;����������У�
			  <SELECT name="city" onChange="location=this.value">
				  <OPTION value="?all=ok">ȫ��
				  </OPTION>
				  <OPTION value="?oneid=0&twoid=0&threeid=0" <%if oneid=0 Then%>selected<%End if%>>��վ</OPTION>
				  <%Set Rsc=Conn.Execute("Select id,twoid,threeid,city from DNJY_city order by id,twoid,threeid")
				  Do While Not Rsc.Eof
					  If Rsc(0)<>0 and Rsc(1)=0 and Rsc(2)=0 Then%>
					  <OPTION style="background-color:#F8F4F5 ;color: #FF0000" value="?oneid=<%=Rsc(0)%>" <%if Rsc(0)=oneid and 0=twoid Then%>selected<%End if%>><%=Rsc(3)%></OPTION>
					  <%end if
					  If Rsc(0)<>0 and Rsc(1)<>0 and Rsc(2)=0 Then%>
					  <OPTION style="background-color:#F8F4F5 ;color: #0000FF" value="?oneid=<%=Rsc(0)%>&twoid=<%=Rsc(1)%>" <%if Rsc(0)=oneid and Rsc(1)=twoid and 0=threeid Then%>selected<%End if%>>��<%=Rsc(3)%></OPTION>
					  <%end if
					  If Rsc(0)<>0 and Rsc(1)<>0 and Rsc(2)<>0 Then%>
					  <OPTION value="?oneid=<%=Rsc(0)%>&twoid=<%=Rsc(1)%>&threeid=<%=Rsc(2)%>" <%if Rsc(0)=oneid and Rsc(1)=twoid and Rsc(2)=threeid Then%>selected<%End if%>>����<%=Rsc(3)%></OPTION>
					  <%End if
					  Rsc.MoveNext
				  Loop
				  %>
              </SELECT>
			   </TD><TD  width="40%" height="20" align="center"> &nbsp;&nbsp;&nbsp;&nbsp;<b>һ��ˢ�����ɹ��JS���룺</b>&nbsp;&nbsp;&nbsp;<a href="createjs_all.asp?city_oneid=<%=rs("city_oneid")%>&city_twoid=<%=rs("city_twoid")%>&city_threeid=<%=rs("city_threeid")%>&city=ok">����������</a> &nbsp;&nbsp;&nbsp;&nbsp; <a href="createjs_all.asp">��ȫվ����</a></TD>
        </TR>
      </TABLE>
    </DIV></TD>
  </TR>  
      <table width='100%' border="0" align=center cellpadding=3 cellspacing=1 bgcolor="#ADAED6">
   
    <% do while not rs.eof
		id=rs("id")
		ADID=rs("ADID")
		ADViews=rs("ADViews")
		ADHits=rs("ADHits")
		ADType=rs("ADType")
		ADSrc=rs("ADSrc")
		ADLink=rs("ADLink")
		ADAlt=rs("ADAlt")
		ADSrc=rs("ADSrc")
		ADWidth=rs("ADWidth")
		ADHeight=rs("ADHeight")
		ADNote=rs("ADNote") 
		ADStopViews=rs("ADStopViews")
		ADStopHits=rs("ADStopHits")
		ADStopDate=rs("ADStopDate")
		IsSys=rs("IsSys")
		adkg=rs("adkg")
ADhost=rs("Advertisement_host")
ADhost=split(ADhost,"|")
ADhost1=ADhost(0)
ADhost2=ADhost(1)
ADhost3=ADhost(2)
ADhost4=ADhost(3)
ADhost5=ADhost(4)
ADhost6=ADhost(5)
ADhost7=ADhost(6)
ADhost8=ADhost(7)
ADhost9=ADhost(8)%>
        <tr bgcolor="#eeeeee">
		<td width="10%">ѡ��</td>
          <td width="10%" height="25"><font color="#0000FF">ID�ţ�<%=ADID%></font> <%if IsStop(ADViews,ADStopViews,ADStopHits,ADHits,ADStopDate) Then Response.Write(" <font color=#FF0000>(�ѹ���)") End If%></td>
          <td width="10%">��ʾ��<%=ADViews%></td>
          <td width="10%">�����<%=ADHits%></td>
		  <td width="15%"><b>���</b><%if adkg=1 then%>��<%else%><font color="#FF0000">��</font><%end if%>&nbsp;&nbsp;&nbsp;&nbsp;<%if adkg=1 then%><a href="createjs.asp?ad=zt&ADID=<%=ADID%>&city_oneid=<%=rs("city_oneid")%>&city_twoid=<%=rs("city_twoid")%>&city_threeid=<%=rs("city_threeid")%>&url=list.asp&page=<%=request("page")%>">��ͣ</a><%else%><a href="createjs.asp?ad=hd&ADID=<%=ADID%>&city_oneid=<%=rs("city_oneid")%>&city_twoid=<%=rs("city_twoid")%>&city_threeid=<%=rs("city_threeid")%>&url=list.asp&page=<%=request("page")%>"><font color="#FF0000">�</font></a><%end if%>&nbsp;&nbsp
          <td width="40%"><b>����</b><a href="edit.asp?id=<%=id%>">�༭</a>&nbsp;&nbsp;<a href="createjs.asp?ADID=<%=ADID%>&city_oneid=<%=rs("city_oneid")%>&city_twoid=<%=rs("city_twoid")%>&city_threeid=<%=rs("city_threeid")%>">ˢ��</a>&nbsp;&nbsp;<a href="../ad_info.asp?ADID=<%=ADID%>&city_oneid=<%=rs("city_oneid")%>&city_twoid=<%=rs("city_twoid")%>&city_threeid=<%=rs("city_threeid")%>" target="_blank">Ԥ��</a>&nbsp;&nbsp;<span id="followImg<%=i%>" style="CURSOR: hand" onClick="loadThreadFollow(<%=i%>,5)">����</span>&nbsp;&nbsp;<span id="followImg<%=i%>" style="CURSOR: hand" onClick="loadThreadFollow(<%=i%>,5)">��ϵ�����</span>
          <%If IsSys = 1 Then
			response.Write "<font color=#FF0000>Ĭ�Ϲ�治��ɾ��</font>"
		Else
		response.Write "<a href='del.asp?id="&id&"' onclick='return delad();'>ɾ��</a>"
		end if%>
          </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="50%" height="25" colspan=2>������ͣ�<%=ShowAdType(ADType,rs("ADSrc"))%></td><td height="25" colspan=2>�����<%=ADWidth%>��<%=ADHeight%></td>
          <td width="50%" colspan=2>��ʾ��ַ��
          <%If ADType<>6 Then%><a href=../<%=ADSrc%> target=_blank><font color="#000000"><%=ADSrc%></a><%else%>����ʾ<%end if%></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td height="25" colspan=4>���ӵ�ַ��<a href=<%=ADLink%> target=_blank><font color="#000000"><%=ADLink%></a></td>
          <td colspan=2>�������ڣ�<%If now()>=ADStopDate then%><font color="#FF0000"><%=ADStopDate%></font><%else%><%=ADStopDate%><%End if%>&nbsp;&nbsp;&nbsp;&nbsp;��ʾ���֣�<%=ADAlt%></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td height="25" colspan=4>����������<%If rs("city_oneid")=0 Then%>��վ<%else%><%=Conn.Execute("select city from DNJY_city where twoid=0 and id="&rs("city_oneid"))(0)%><%If rs("city_twoid")>0 then%>/<%=Conn.Execute("select city from DNJY_city where id="&rs("city_oneid")&" and threeid=0 and twoid="&rs("city_twoid")&"")(0)%><%End If%><% If rs("city_threeid")>0 then%>/<%=Conn.Execute("select city from DNJY_city where id="&rs("city_oneid")&" and twoid="&rs("city_twoid")&" and threeid="&rs("city_threeid"))(0)%><%End if%><%End if%></td>
          <td colspan=2>��    ע��<%=ADNote%></td>                         
        </tr>
          <tr class="jg<%=skin%>"> 
    <td height="5" colspan=8></td>
  </tr>
  <tr style="display:none" id="follow<%=i%>">
    <td width="100%" height="24" colspan="11" style="border-bottom-style: solid; border-bottom-width: 1" bgcolor="#FFF8EC">
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#6699cc">

  <tr>
    <td height="120" align="center">
<TABLE>
  <tr><td height="28" colspan="2"><div align="center" ><b>���[<%=ADID%>]����</b></div></td> <td height="28"><div align="center" ><b><b>�����������ϵ��Ϣ</b></b></div></td> 
  </tr>
  <TR>
	<TD width="250"><textarea name="S1" cols="50" rows="6" class="input2"><!--JS�̶��������÷�ʽ����JS�ļ��޷���ʾ������̬ҳ������ã����뿪ʼ-->
  <script src="ads/JS/<%=ADID%>_<%=rs("city_oneid")%>_<%=rs("city_twoid")%>_<%=rs("city_threeid")%>.js"></script>
  <!--���������-->
      </textarea>
        <textarea name="S1" cols="50" rows="6" class="input2"><!--JS��̬�������÷�ʽ����JS�ļ�����վ�����棬ֻ�ܶ�̬ҳ���ã����뿪ʼ-->
  <script src="&lt;%=adjs_path("ads/js","<%=ADID%>",c1&"_"&c2&"_"&c3)%>"></script>
  <!--���������-->
      </textarea> </TD>
	<TD width="250">      <textarea name="textarea" cols="50" rows="6" class="input2"><!--ASP�̶��������÷�ʽ���루ֻ�ܶ�̬ҳ���ã���ʼ-->
  <script src="ads/ad.asp?adid=<%=ADID%>&city_oneid=<%=rs("city_oneid")%>&city_twoid=<%=rs("city_twoid")%>&city_threeid=<%=rs("city_threeid")%>"></script>
<!--���������-->
      </textarea>
      <textarea name="textarea" cols="50" rows="6" class="input2"><!--ASP��̬�������÷�ʽ���루ֻ�ܶ�̬ҳ���ã���ʼ-->
  <script src="ads/ad.asp?adid=<%=ADID%>&city_oneid=&lt;%=c1%>&city_twoid=&lt;%=c2%>&city_threeid=&lt;%=c3%>"></script>
<!--���������-->
      </textarea></TD>
	<TD width="250">
	����:<%=ADhost1%><br>�绰:<%=ADhost2%><br>�ֻ�:<%=ADhost3%><br>Q &nbsp;Q:<%=ADhost4%><br>����:<%=ADhost5%><br>��ַ:<%=ADhost6%><br>�ʱ�:<%=ADhost7%><br>����:<%=ADhost8%> Ԫ<br>��ע:<%=ADhost9%></TD>
  </TR>
  </TABLE> 
  
	  </td> 
  </tr>

</table>
    </td>
  </tr>
<%i=i+1
if i>=MaxPerPage then Exit Do
rs.movenext
loop%>
</table>
<%
Call CloseRs(rs)
Call CloseDB(conn)
End Sub%>  
<%'=============���÷�ҳ��ʾ��ʼ 
  Function showpage(totalnumber,maxperpage,filename)  
      Dim n      
    If totalnumber Mod maxperpage=0 Then  
     n= totalnumber \ maxperpage  
    Else
     n= totalnumber \ maxperpage+1  
    End If %>
    <form method=Post action=<%=filename%>>
    <p align="center"> 
<%If CurrentPage<2 Then  %>
    �� ҳ ��һҳ
    <% Else  %>
    <a href=<% = filename %>?page=1&all=<%=Request("all")%>&<%=Temp%>>�� ҳ</a>
    <a href=<% = filename %>?page=<% = CurrentPage-1 %>&all=<%=Request("all")%>&<%=Temp%>>��һҳ</a> 
    <% End If 
    If n-currentpage<1 Then  %>
    ��һҳ β ҳ
    <%  Else  %>
    <a href=<% = filename %>?page=<% = (CurrentPage+1) %>&all=<%=Request("all")%>&<%=Temp%>>��һҳ</a> 
    <a href=<% = filename %>?page=<% = n %>&all=<%=Request("all")%>&<%=Temp%>>β ҳ</a>&nbsp;&nbsp;
    <% End If  %>
 ҳ�Σ�<b><font color=red><% = CurrentPage %></font></b>/<b><% = n %></b>ҳ <b><%=maxperpage%></b>����¼/ҳ  ��<b><%=totalnumber %></b>����¼    
ת����<select name="cndok" onchange="javascript:location=this.options[this.selectedIndex].value;">
<%
for i = 1 to n
if i = CurrentPage then%>
<option value="<% = filename %>?page=<%=i%>&all=<%=Request("all")%>&<%=Temp%>" selected>��<%=i%>ҳ</option>  
<%else%>
<option value="<% = filename %>?page=<%=i%>&all=<%=Request("all")%>&<%=Temp%>">��<%=i%>ҳ</option>  
<%
end if
next
%>
</select></font>
</form>
<%End Function 
'=============���÷�ҳ��ʾ����%>
</body>
</html>

<%'����Ƿ����
function IsStop(ADViews,ADStopViews,ADStopHits,ADHits,ADStopDate)
	IsStop=false
	If ( ADStopViews <> 0 and ADViews > ADStopViews) Then 
		IsStop=true
		Exit function
	ElseIf ( ADStopHits <> 0 and ADHits > ADStopHits) Then
		IsStop=true
		Exit function
	ElseIf ( DateDiff("d",Now(),ADStopDate)<1 ) Then	
		IsStop=true
		Exit function
	End If
end function

'�жϹ������
function ShowAdType(ADType,ADSrc)
	Dim ADExt
	ADExt="ͼƬ"
	If InStr(1,ADSrc,".swf",1)>0 Then ADExt="FLASH"
	Select Case ADType
		Case 1
			ShowAdType="��ͨ"&ADExt
		Case 2
			ShowAdType="ȫ������"&ADExt
		Case 3
			ShowAdType="���¸��� - ��"&ADExt
		Case 4
			ShowAdType="���¸��� - ��"&ADExt
		Case 5
			ShowAdType="������ʧ"&ADExt
		Case 6
			ShowAdType="��ҳ�Ի���"
		Case 7
			ShowAdType="�ƶ�͸���Ի���"
		Case 8
			ShowAdType="���´���"
		Case 9
			ShowAdType="�����´���"
		Case 10 
			ShowAdType="����ʽ��棨�̶���"
        Case 11 
			ShowAdType="����ʽ��棨������"	
        Case 12 
			ShowAdType="��������ֹ��"
        Case 13 
			ShowAdType="��ͨ���ֹ��"
        Case 14 
			ShowAdType="��д������"
        Case 15 
			ShowAdType="FlashͼƬ���"
        Case 16 
			ShowAdType="��Ƶ���"				
		Case else
			ShowAdType="<font color=red><b>����!��������ȷ��ʾ</b>"
	End Select
end function
%>