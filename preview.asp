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
<link rel="stylesheet" type="text/css" href="1.CSS">
</head>
<%
dim rsper,k,sdname,sdkg,sdfg,mpname,mpkg,mpfg,vip,guest,permissions
username=trim(request("username"))
Readinfo=trim(request("Readinfo"))
If trim(request("Readinfo"))="" Then Readinfo=0
Call OpenConn
set rsper = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_user] where username='"&username&"'" 
rsper.open sql,conn,1,1
if rsper.eof then
response.write "<p><center><li>�ۣ��������Ҵ�ط�ඣ��û�û��ע����ô�죡"
response.end
end if
vip=rsper("vip")
sdkg=rsper("sdkg")
sdfg=rsper("sdfg")
sdname=rsper("sdname")
mpkg=rsper("mpkg")
mpfg=rsper("mpfg")
mpname=rsper("mpname")

If request.cookies("dick")("username")<>username Then'�Լ�������ȫȨ�Ķ�
guest=0'��¼��Ա
If request.cookies("dick")("username")<>"" Then guest=conn.Execute("Select vip From DNJY_user Where username='"&request.cookies("dick")("username")&"'")(0)
If Readinfo=1 And request.cookies("dick")("username")="" Then
permissions="��ͨ��ԱȨ���Ķ�"
End If
If Readinfo=2 And guest<1 Then
permissions="VIP��ԱȨ���Ķ�"
End If
End If%>
<body topmargin="0" background="images/bg1.gif">
<div align="center">

</div>
<table cellSpacing="0" cellPadding="0" width="980" align="center" bgColor="#ffffff" border="0" height="410">
  <tr>
    <td width="980" height="331">
    <table cellSpacing="0" cellPadding="0" width="980" border="0" height="406">
      <tr>
        <td vAlign="top" width="210" height="406">
        <table cellSpacing="0" cellPadding="0" width="97%" border="0">
                   <tr>
            <td>
           <!--#include file=left.asp--></td>
          </tr>
        </table>
        </td>
		<td width="5">&nbsp;</td>
        <td vAlign="top" width="760" height="406">
        <div align="center">        
        <table border="0" cellpadding="0" cellspacing="0"   width="100%" height="222">
		  <table width="760" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="ty1">
          <tr>
            <td width="100%" height="50" background="images/user_1.gif" colspan="6">��
              <div align="center"><b>��Ա��ϸ��Ϣ�������ײο�</b></div></td>
          </tr>
<tr>
<td>
<table width="750" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" align="center">
<TR>
	<TD valign="top" width="450"><TABLE>
	<tr>
    <td align="right">��ԱID�ţ�</td><td><%=rsper("username")%></td>
    </tr>
	<tr>
    <td align="right">��Ա������</td><td><%if Readinfo>0 then%><%=permissions%><%else%><%=rsper("name")%>&nbsp;<%if rsper("sex")=1 then%>����<%else%>Ůʿ<%end if%><%end if%></td>
    </tr>
	<tr>
    <td align="right">VIP ��Ա��</td><td><%if rsper("vip")=1 then%>��<%else%>��<%end if%>&nbsp;&nbsp;&nbsp;&nbsp;<%if sdkg=1 And sdname<>"" then%> <a  href="user/com.asp?username=<%=rsper("username")%>&com=<%=rsper("username")%>&<%=C_ID%>" target="_blank"><img src="img/dian.gif" title="�˻�Ա�ѿ��е���" ></a><%end if%> &nbsp;&nbsp;&nbsp;<%if mpkg=1 And mpname<>"" then%><a  href="company.asp?username=<%=rsper("username")%>&<%=C_ID%>" target="_blank"><img src="img/qy.gif" title="�˻�Ա�ѿ�����ҵ��Ƭ" ></a><%end if%></td>
    </tr>
	<tr>
    <td align="right">ע��ʱ�䣺</td><td><%=rsper("zcdata")%></td>
    </tr>
	<tr>
    <td align="right">վ�����ۣ�</td><td><img src="img/hp.jpg"> ���ż�¼ <%=rsper("goodfaith")%> ��&nbsp;&nbsp;| <img src="img/cp.jpg"> ʧ�ż�¼ <%=rsper("bpromises")%> ��</td>
    </tr>
	<tr>
    <td align="right" valign="top">���Ѵ�֣�</td><td>

<%Dim df,cs,pj
df=rsper("wevaluation")
cs=rsper("participants")
If rsper("wevaluation")=0 Then df=5
If rsper("participants")=0 Then cs=1
pj=Formatnumber(df/(cs*5)*100,0,0,0,0)%>
	<style type="text/css"> 
body{ 
text-align:center; 
} 
.graph{ 
width:355px; 
border:1px solid #F8B3D0; 
height:15px; 
} 
#bar{ 
display:block; 
background:#FFE7F4; 
float:left; 
height:100%; 
text-align:center; 
} 
#barNum{ 
position:absolute; 
} 
</style> 
<script type="text/javascript"> 
function $(obj){ 
return document.getElementById(obj); 
} 
function go(){ 
$("bar").style.width = parseInt($("bar").style.width) + 1 + "%"; 
$("bar").innerHTML = $("bar").style.width; 
if($("bar").style.width =="<%=pj%>%"){ 
window.clearInterval(bar); 
} 

} 
var bar = window.setInterval("go()",30); 
window.onload = function(){ 
bar; 
} 
</script> 
<div class="graph"> 
<strong id="bar" style="width:1%;"></strong> 
</div>
<img src="img/pj.jpg" width="370"><br>���Ѳ��� <%=cs%> �˴�&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ƽ���÷� <%=Formatnumber(df/cs,1,0,0,0)%> ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<INPUT TYPE=button VALUE="��Ҫ����" ONCLICK="window.open('user/evaluation.asp?username=<%=rsper("username")%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbarsper=yes,scrollbars=yes,resizable=no,copyhistory=no,width=400,height=300,left=300,top=100')" style="color: #0060C5; background-color: #E1E2DC">	
	</td>
    </tr>

	</TABLE>		  
	</TD>
	<TD valign="top"><TABLE>
	<tr>
    <td align="right">��ѶQQ�ţ�</td><td><%if Readinfo>0 then%><%=permissions%><%else%><%=rsper("qq")%><%end if%></td>
    </tr>
	<tr>
    <td align="right">��ϵ�绰��</td><td><%if Readinfo>0 then%><%=permissions%><%else%><%=rsper("dianhua")%><%end if%></td>
    </tr>
	<tr>
    <td align="right">��ϵ�ֻ���</td><td><%if Readinfo>0 then%><%=permissions%><%else%><%=rsper("CompPhone")%><%end if%></td>
    </tr>
	<tr>
    <td align="right">վ�ڶ��ţ�</td><td><%if Readinfo>0 then%><%=permissions%><%else%><a href='messh.asp?username2=<%=trim(rsper("username"))%>&<%=C_ID%>'>��վ�ڶ��Ÿ���</a><%end if%></td>
    </tr>
	<tr>
    <td align="right">�������䣺</td><td><%if Readinfo>0 then%><%=permissions%><%else%><INPUT TYPE=button VALUE="���ʼ�����" ONCLICK="window.open('user/user_mail.asp?username=<%=rsper("username")%>&email=<%=rsper("email")%>&mailzt=������Ϣ','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbarsper=yes,scrollbars=yes,resizable=no,copyhistory=no,width=650,height=300,left=300,top=100')" style="color: #666666; background-color: #E1E2DC"><%end if%></td>
    </tr>
	<tr>
    <td align="right">��˾���ƣ�</td><td><%if Readinfo>0 then%><%=permissions%><%else%><%end if%></td>
    </tr>
	<tr>
    <td align="right">�������룺</td><td><%if Readinfo>0 then%><%=permissions%><%else%><%=rsper("youbian")%><%end if%></td>
    </tr>
	<tr>
    <td align="right">ͨ�ŵ�ַ��</td><td><%if Readinfo>0 then%><%=permissions%><%else%><%=rsper("dizhi")%><%end if%></td>
    </tr>
	</TABLE>
	</TD>
</TR>
</TABLE>
</td>
</tr>          
           <tr>
            <td width="100%" height="12" colspan="6"></td>
          </tr>
            <td width="100%" height="12" colspan="6"><div align="center"><font color="#ff0000"><%=rsper("name")%>(<%=rsper("username")%>)</font>��������Ϣ<%If rsper("vip")=0 then%> �����û��ѷ���<%=rsper("xxts")%>����Ϣ��������VIP��Ա�����ֻ��ʾ����<%=huiyuanxx%>����Ϣ��<%end if%></div></td>          
          <%rsper.close%>          
            <td width="100%" height="24" background="images/user_2.gif" colspan="6"><a name="1"></a></td>
          </tr>
          <tr>
            <td width="100%" height="24" colspan="6">
            <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="99%" height="100%">
<%
dim ThisPage,Pagesize,Allrecord,Allpage,b,bb
k=0
if request("page")="" then
  ThisPage=1		
else
  ThisPage=request("page")
end If
if request.cookies("administrator")<>"" then
if vip=1 then
sql = "select id,biaoti,a,b,c,fbsj,llcs,hfcs,Readinfo from DNJY_info where username='"&username&"' order by b "&DN_OrderType&",id "&DN_OrderType&""
else
sql = "select top "&huiyuanxx&" id,biaoti,a,b,c,fbsj,llcs,hfcs,Readinfo from DNJY_info where username='"&username&"' order by b "&DN_OrderType&",id "&DN_OrderType&""
end If
Else
if vip=1 then
sql = "select id,biaoti,a,b,c,fbsj,llcs,hfcs,Readinfo from DNJY_info where yz=1 and username='"&username&"' order by b "&DN_OrderType&",id "&DN_OrderType&""
else
sql = "select top "&huiyuanxx&" id,biaoti,a,b,c,fbsj,llcs,hfcs,Readinfo from DNJY_info where yz=1 and username='"&username&"' order by b "&DN_OrderType&",id "&DN_OrderType&""
end If
End if
rsper.open sql,conn,1,1
if rsper.eof then
response.write "���û���û�з�����Ϣ��"
else
rsper.Pagesize=10
Pagesize=rsper.Pagesize
Allrecord=rsper.Recordcount
Allpage=rsper.Pagecount
if ThisPage<1 then                           
ThisPage=1
end if
'On Error Resume Next
rsper.move (ThisPage-1)*Pagesize
%>
              <tr>
                <td width="7%" align="center" height="25" style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1" bordercolor="#F4F4F2">���</td>
                <td width="66%" align="center" height="25" style="border-top-style: solid; border-top-width: 1" bordercolor="#F4F4F2">&nbsp; ��&nbsp;&nbsp; 
                Ϣ&nbsp;&nbsp; ��&nbsp;&nbsp; ��</td>
                <td width="14%" align="center" height="25" style="border-top-style: solid; border-top-width: 1" bordercolor="#F4F4F2">
                ��/��</td>
                <td width="13%" align="center" height="25" style="border-right-style: solid; border-right-width: 1; border-top-style: solid; border-top-width: 1" bordercolor="#F4F4F2">����ʱ��</td>
              </tr>
<%
do while not rsper.eof
b=trim(rsper("b"))
bb=len(b)
%>

              <tr>
                <td width="7%" align="center" height="25" style="border-left-style: solid; border-left-width: 1" bordercolor="#F4F4F2"><%=k+1%></td>
                <td width="66%" align="center" height="25" style="border-bottom-style: none; border-bottom-width: medium" bordercolor="#F4F4F2"><p align="left"><%if rsper("c")=1 then%><img src="images/num/pic.gif"><%end if%><a target="_blank" href="<%=x_path("",rsper("id"),"",rsper("Readinfo"))%>"><%if rsper("a")="0" then%><%=rsper("biaoti")%><%else%><font color="#<%=rsper("a")%>"><%=rsper("biaoti")%></font><%end if%></a><%if b<>0 then%><img src="images/num/jsq.gif"><%for i=1 to bb%><img src="images/num/<%=Mid(b,i,1)%>.gif"><%next%><%end if%></td>
                <td width="14%" align="center" height="25" style="border-bottom-style: none; border-bottom-width: medium" bordercolor="#F4F4F2"><%=rsper("llcs")%>/<%=rsper("hfcs")%></td>
                <td width="13%" align="center" height="25" style="border-right-style: solid; border-right-width: 1; border-bottom-style: none; border-bottom-width: medium" bordercolor="#F4F4F2"><%=datevalue(rsper("fbsj"))%></td>
              </tr>
<%
k=k+1
rsper.movenext
if k>=Pagesize then exit do
loop
%>
              <tr>
                <td width="100%" align="center" height="13" colspan="4" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1" bordercolor="#F4F4F2">
<%
if Allrecord>rsper.Pagesize then
call pages()
end if
sub pages()
%>
<table border="0" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="569">
<tr> 
<td height="24" width="569" colspan="4">
<p align="center">
��</td>
</tr>
<tr> 
<td height="24" width="151" align="center">
����&nbsp;<font color="#CC5200"><%=Allrecord%></font>&nbsp;����¼</td>
<td height="24" width="126" align="center">
�� <font color="#CC5200"><%=Allpage%></font> ҳ</td>
<td height="24" width="118" align="center">
�����ǵ� 
                <font color="#CC5200"><%=ThisPage%></font> ҳ</td>
<td height="24" width="174" align="center">
<%
if ThisPage<2 then     
response.write "<font color=""#808080"">��ҳ</font>&nbsp;"
response.write "<font color=""#808080"">��һҳ</font>&nbsp;"     
else     
response.write "<a href=?page=1&username="&username&"#1>��ҳ</a>&nbsp;"
response.write "<a href=?page="&ThisPage-1&"&username="&username&"#1>��һҳ</a>&nbsp;"     
end if
if Allpage-ThisPage<1 then     
response.write "<font color=""#808080"">��һҳ</font>&nbsp;"
response.write "<font color=""#808080"">βҳ</font>&nbsp;"  
else     
response.write "<a href=?page="&(ThisPage+1)&"&username="&username&"#1>��һҳ</a>&nbsp;"   
response.write "<a href=?page="&Allpage&"&username="&username&"#1>βҳ</a>&nbsp;"     
end if
%></td>
</tr>
</table>
<%end sub%>
                </td>
              </tr>
              <tr>
                <td width="100%" align="center" height="12" colspan="4" style="border-left-style: none; border-left-width: medium; border-right-style: none; border-right-width: medium; border-top-style: solid; border-top-width: 1; border-bottom-style: none; border-bottom-width: medium" bordercolor="#F4F4F2">
                </td>
              </tr>
<%end if%>
            </table>
            </td>
          </tr>
        </table>
          </center>
        </div>
        </td>
      </tr>
    </table>
    </td>
  </tr>
</table>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="772" height="23">
    <tr>
      <td width="100%" height="16" bgcolor="#FFFFFF"></td>
    </tr>
    <tr>
      <td width="100%" height="7" bgcolor="#FFFFFF"> </td>
    </tr>
    </table>
  </center>
</div>
</body><%rsper.close
set rsper=nothing%>
<!--#include file=footer.asp-->
</html>
