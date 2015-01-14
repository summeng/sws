<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("12")%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/style.css" type="text/css">
<BODY background="img/back.gif">

<script language=javascript>
<!--
function left_menu(meval)
{
  var left_n=eval(meval);
  if (left_n.style.display=="none")
  { eval(meval+".style.display='';"); }
  else
  { eval(meval+".style.display='none';"); }
}
function DoEmpty(params)
{
if (confirm("真的要删除这条记录吗？删除后此记录里的所有内容都将被删除并且无法恢复！"))
window.location = params ;
}
function mm_jumpmenu(targ,selobj,restore){ //v3.0
  eval(targ+".location='"+selobj.options[selobj.selectedindex].value+"'");
  if (restore) selobj.selectedindex=0;
}
//document.all.left_sys.style.display='';
//document.all.left_bm.style.display='none'
-->
</script>
<script language="JavaScript">
<!--
function CheckAll(form)
{
for (var i=0;i<form.elements.length;i++)
{
var e = form.elements[i];
if (e.name != 'chkall')
e.checked = form.chkall.checked;
}
}

//-->
</script>
<style type="text/css">
<!--
body {
	background-color: #E7EEF5;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>

<title>管理都市114</title>
</head>
<body ><br>
<table width="90%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolorlight="#335078" bordercolordark="#FFFFFF" bgcolor="#C5D5E4"  style="font-size:12px; letter-spacing:1px;">
  <tr> 
    <td height="20"  colspan="15" bgcolor="#C5D5E4" align="center"><span style="color: #FF0000"><strong><font style="font-size:14px">都市114管理</font></strong></span></td>
  </tr>


          <tr> 
 <td  bgcolor="#D9E4EE" colspan="5">
 <%
 dim key,follows,d_name,co,d_tel,hfje,xsts,updatetime,dqsj,known,dick,dd,selectkey,selectm,alldel,Num,u,page,oneid
	city_oneid=Request("city_one")
	city_twoid=Request("city_two")
	city_threeid=Request("city_three")
   If city_oneid="" Then city_oneid=0
   If city_twoid="" Then city_twoid=0
   If city_threeid="" Then city_threeid=0
   dick=trim(request("dick"))
   If dick="" Then dick=0
 key=Request("key")
 selectkey=trim(request("selectkey"))
selectm=trim(request("selectm"))
if selectkey="" then
selectkey=request.QueryString("selectkey")
end if
if selectm="" then
selectm=request.QueryString("selectm")
end if
id=trim(request("id"))

If key="del" then
if not isnumeric(id) or id="" then
response.write "<li>参数错误！"
response.end
end if
set rs=server.createobject("adodb.recordset")
sql="delete from [DNJY_114] where id="&cstr(id)
rs.open sql,conn,1,3
response.Write "<script language=javascript>alert('删除成功!');location.replace('?page="&page&"&dick="&dick&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&"');</script>"

end if

If key="kg" then
known=trim(request("known"))
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_114] where id="&trim(id)
rs.open sql,conn,1,3
If known="no" Then
rs("known")=1
Else
rs("known")=0
End if
rs.update
Call CloseRs(rs)
end If

%>
 <form method="POST" name="form" id="form" language="javascript"  action="?">
<% set rs=server.createobject("adodb.recordset")
set rs=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")%>
<SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
	dim count5:count5 = 0
        do while not rs.eof 
        %>
subcat[<%=count5%>] = new Array("<%=rs("city")%>","<%=rs("id")%>","<%=rs("twoid")%>");
        <%
        count5 = count5 + 1
        rs.movenext
        loop
Call CloseRs(rs)
        %>
onecount=<%=count5%>;
          </SCRIPT>
          <%
set rs=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid>0")
%>
          <SCRIPT language = "JavaScript">
var onecount4;
onecount4=0;
subcat4 = new Array();
        <%
		dim count4:count4 = 0
        do while not rs.eof 
        %>
subcat4[<%=count4%>] = new Array("<%=rs("city")%>","<%=rs("id")%>","<%=rs("twoid")%>","<%=rs("threeid")%>");
        <%
        count4 = count4 + 1
        rs.movenext
        loop
Call CloseRs(rs)
        %>
onecount4=<%=count4%>;

function changelocation(locationid)
    {
    document.form.city_two.length = 0; 
	document.form.city_two.options[0] = new Option('选择城市','');
	document.form.city_three.length = 0; 
	document.form.city_three.options[0] = new Option('选择城市','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.form.city_two.options[document.form.city_two.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
	
	function changelocation4(locationid,locationid1)
    {
    document.form.city_three.length = 0; 
    document.form.city_three.options[0] = new Option('选择城市','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount4; i++)
        {
            if (subcat4[i][2] == locationid)
            { 
			if (subcat4[i][1] == locationid1)
			{
                document.form.city_three.options[document.form.city_three.length] = new Option(subcat4[i][0], subcat4[i][3]);
            }        
        }
       } 
    }            
              </SCRIPT>
          <SELECT name="city_one" size="1" id="select" onChange="changelocation(document.form.city_one.options[document.form.city_one.selectedIndex].value)">
            <OPTION value="" selected>选择地区</OPTION>
            <%set rs=conn.execute("select * from DNJY_city where id>0 and twoid=0  order by indexid "&DN_OrderType&"")
if rs.eof or rs.bof then
response.write "<option value=''>没有分类</option>"
else
do until rs.eof
response.write "<option value='"&rs("id")&"'>"&rs("city")&"</option>"
    rs.movenext
    loop
end if
Call CloseRs(rs)
%>
          </SELECT>
          <SELECT name="city_two" id="select6" onChange="changelocation4(document.form.city_two.options[document.form.city_two.selectedIndex].value,document.form.city_one.options[document.form.city_one.selectedIndex].value)">
            <OPTION value="" selected>选择城市</OPTION>
          </SELECT>
          <SELECT name="city_three" id="select7">
            <OPTION value="" selected>选择城市</OPTION>
          </SELECT>
 <input class="inputb" name="I2" type="submit" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #a2a2a2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #a2a2a2 1px solid; BORDER-TOP: #ffffff 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;"  value="查询" border="0" /></font>
</form></td> <td  bgcolor="#D9E4EE" colspan="5"><a href="?dick=0&page=<%=Page%>&city_one=0&city_two=0&city_three=0">全部地区114</a>&nbsp;&nbsp;&nbsp;<a href="?dick=0&page=<%=Page%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>">本地区114</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="?dick=1&page=<%=Page%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&selectkey=<%=selectkey%>&selectm=<%=selectm%>">服务到期的都市114</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="?dick=2&page=<%=Page%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&selectkey=<%=selectkey%>&selectm=<%=selectm%>">隐藏的都市114</a>&nbsp;&nbsp;&nbsp;<a href="#" onClick="window.open('admin_114edit.asp?admin114=2','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=600,height=350,left=300,top=100')">添加</a></td>      
          </tr>
<form name="form1" method="post" action="admin_114edit.asp?action=del">

  <tr> 
    <td width="3%" height="25" align="center" bgcolor="#cccccc">选择</td>
    <td width="5%" align="center" bgcolor="#cccccc">编号</td>
    <td width="15%" align="center" bgcolor="#cccccc">服务项目</td>
    <td width="15%" align="center" bgcolor="#cccccc">电话</td>
    <td width="8%" align="center" bgcolor="#cccccc">城市</td>
    <td width="5%" align="center" bgcolor="#cccccc">座位票价(元)</td>
    <td width="5%" align="center" bgcolor="#cccccc">显示期限(天)</td>
    <td width="25%" align="center" bgcolor="#cccccc">添加/到期时间</td>
    <td width="5%" align="center" bgcolor="#cccccc">活动</td>
    <td width="10%" align="center" bgcolor="#cccccc">操作</td>
  </tr>
<%
    set rs=server.CreateObject("adodb.recordset")
	IF city_oneid=0 then
	ttcc114=""
	elseIF city_threeid>0 then
	ttcc114=" city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid
	elseIF city_twoid>0 then
	ttcc114=" city_oneid="&city_oneid&" and city_twoid="&city_twoid
	Else
	ttcc114=" city_oneid="&city_oneid
	End If

sql = "select d_name,d_tel,id,hfje,xsts,updatetime,dqsj,known,co,city_one,city_two,city_three from DNJY_114 where d_name<>''"
If city_oneid>0 Then sql=sql&" and "&ttcc114&""

Select Case dick
Case "1"
sql=sql&" and dqsj<"&DN_Now&""
Case "2"
sql=sql&" and known=1"
End Select

			select case selectm
			case ""
            sql=sql
		    case "name"
			sql=sql&" and d_name like '%"&selectkey&"%'"
			case "dianhua"
            sql=sql&" and d_tel like '%"&selectkey&"%'"
            end Select
sql=sql&" order by id "&DN_OrderType&""           
rs.open sql,conn,1,1

	  
		  if rs.EOF and rs.BOF Then
		  IF oneid>0 Then dd=conn.Execute("Select city From DNJY_city Where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid&"")(0)
          response.write"<tr bgcolor=#FFFFFF><td colspan='10'><p align='center'><font color='red'>暂无"&dd&"都市114服务！</font></td></tr></table><br>"
		  follows=0
		  Else
Dim Pagesize,Allrecord,Allpage,ThisPage,k
if request("page")="" then
  ThisPage=1		
else
  ThisPage=request("page")
end if
rs.Pagesize=20 '每页显示条数
Pagesize=rs.Pagesize
Allrecord=rs.Recordcount
Allpage=rs.Pagecount
if ThisPage<1 then                           
ThisPage=1
end if
On Error Resume Next
rs.move (ThisPage-1)*Pagesize
k=1
		  do while not rs.EOF
		  i=i+1
%>
  <tr bgcolor="#ffffff">
   <td height="22" align="center" bgcolor="#F4F7F9"><input name="DelID" type="checkbox" id="DelID" value="<%=rs("id")%>"></td>
    <td height="22" align="center" bgcolor="#F4F7F9"><%=i%></td>
    <td align="center" bgcolor="#F4F7F9"><%if rs("co")<>"0" Then%><font color=#<%=rs("co")%>><%=trim(rs(0))%><%else%><%=trim(rs(0))%><%End if%></td>
    <td bgcolor="#F4F7F9"><%=trim(rs(1))%></td>
     <TD align="center" bgcolor="#F4F7F9"><A title=<%=rs("city_one")%>/<%=rs("city_two")%>/<%=rs("city_three")%> href="#"><%=rs("city_one")%></A></TD>

    <td align="center" bgcolor="#F4F7F9"><%=rs(3)%></td>
    <td align="center" bgcolor="#F4F7F9"><%=rs(4)%></td>
    <td align="center" bgcolor="#F4F7F9"><%=(rs(5))%><font color="#FF0000">/</font><%If now()>=rs(6) then%><font color="#FF0000"><%=rs(6)%></font><%else%><%=rs(6)%><%End if%></td>

    <td align="center" bgcolor="#F4F7F9"><%if rs("known")=0 then%><a href="?known=no&id=<%=rs("id")%>&key=kg&page=<%=ThisPage%>&dick=<%=dick%>&city_one=<%=city_oneid%>&city_two=<%city_twoid%>&city_three=<%city_threeid%>">√</a><%else%><b><a href="?known=yes&id=<%=rs("id")%>&key=kg&page=<%=ThisPage%>&dick=<%=dick%>&city_one=<%=city_oneid%>&city_two=<%city_twoid%>&city_three=<%city_threeid%>"><font color="#FF0000">ⅹ</font></a></b><%end if%></td>
	<td align="center" bgcolor="#F4F7F9"><a href="#" onClick="window.open('admin_114edit.asp?admin114=1&id=<%=rs("id")%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=600,height=350,left=300,top=100')">修改</a> 
	<%if session("aleave")="check" then%>删除<%else%><a href="javascript:DoEmpty('?key=del&id=<%=rs("id")%>&page=<%=ThisPage%>&dick=<%=dick%>&city_one=<%=city_oneid%>&city_two=<%city_twoid%>&city_three=<%city_threeid%>');">删除</a><%end if%></td>
  </tr>
 
       <%rs.MoveNext
	   if k>=Pagesize then exit Do
	   k=k+1
          loop
          follows=rs.RecordCount
          end If
          Call CloseRs(rs)
		  conn.close:Set conn=nothing%>
</table>
  <p align="center">

    <input type="checkbox" name="chkall" onClick="CheckAll(this.form)">
    全部选中 
    <input type="submit" name="Submit" value="删除所选114" onClick="{if(confirm('该操作不可恢复！\n\n确实删除选定的文章？')){return true;}return false;}"> 
  </p>
</form>
 <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#6699cc">
<tr> 
<td width="595" height="25" align="center" bgcolor="#D9E4EE">
  共有&nbsp;<font color="#CC5200"><%=Allrecord%></font>&nbsp;条记录 共 <font color="#CC5200"><%=Allpage%></font> 页 现在是第 <font color="#CC5200"><%=ThisPage%></font> 页
<%
if ThisPage<2 then     
response.write "<font color=""#808080"">首页</font>&nbsp;"
response.write "<font color=""#808080"">上一页</font>&nbsp;"     
else     
response.write "<a href=?page=1&dick="&dick&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&"&selectkey="&selectkey&"&selectm="&selectm&">首页</a>&nbsp;"
response.write "<a href=?page="&ThisPage-1&"&dick="&dick&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&"&selectkey="&selectkey&"&selectm="&selectm&">上一页</a>&nbsp;"     
end if
if Allpage-ThisPage<1 then     
response.write "<font color=""#808080"">下一页</font>&nbsp;"
response.write "<font color=""#808080"">尾页</font>&nbsp;"  
else     
response.write "<a href=?page="&(ThisPage+1)&"&dick="&dick&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&"&selectkey="&selectkey&"&selectm="&selectm&">下一页</a>&nbsp;"   
response.write "<a href=?page="&Allpage&"&dick="&dick&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&"&selectkey="&selectkey&"&selectm="&selectm&">尾页</a>&nbsp;"     
end if
%></td>
</tr>
      </table>   

<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#6699cc">
  <tr> 
    <td align="center" colspan="3">搜索查询</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <form name="form2" method="post" action="?page=<%=Page%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>"&selectkey=<%=selectkey%>&selectm=<%=selectm%>">
      <td>
	  <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
          <tr bgcolor="#FFFFFF" align="center"> 
            <td width="35%"> <input name="selectkey" type="text" id="selectkey" onFocus="this.value=''" value="请输入关健字" size="30"> 
             <select name="selectm" id="select">                
                <OPTION VALUE="name">按服务名称</OPTION> 
                <OPTION VALUE="dianhua">按电话</OPTION>
              </select> <input type="submit" name="Submit" value="查 询" ></td>
          </tr>
        </table></td>
    </form>
  </tr>
</table>

</body>
</html>
