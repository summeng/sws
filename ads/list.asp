<!--#include file="../conn.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
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
if (confirm("删除广告将同时删除相关图片和JS代码，不可以恢复，确认删除吗？")){return true;}
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
<!---弹出面板开始--->
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
<!---弹出面板结束--->
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

'监测是否存在错误
If err.number <> 0 Then
	Response.Write "打开数据库出错,请检查conn.asp中数据库地址是否正确"
	Response.End
End If

If rs.bof or rs.eof Then
	rs.Close
	if Request.QueryString("action")="stop" Then
	%>
	<table width="100%" height="200" border="0" align="center" cellpadding="0" cellspacing="1">
  <tr>
    <td align="center"><font color="#990000">暂时没有过期广告！</font></td>
  </tr>
</table>
	<%Else%>
	<table width="100%" height="200" border="0" align="center" cellpadding="0" cellspacing="1">
  <tr>
    <td align="center" bgcolor="#DEDBEF">暂时没有符合条件广告！-- <a href=add.asp>立即<font color="#0000ff">添加新广告</font></a></td>
  </tr>
</table>
	<%
	End If
	Response.End
End If

%>
<%'=============分页定义开始，可放在数据库打开前或后
                 dim action
    action=request.QueryString("action")   
    Const MaxPerPage=10   '定义每页显示记录数，可根据实际自定义
       dim totalPut   
       dim CurrentPage
       dim TotalPages
       dim j

        if Not isempty(request("page")) And request("page")<>"" then
          currentPage=Cint(request("page"))
       else
          currentPage=1
       end if        
'=============分页定义结束%>
<%'=============分页类代码开始，需放在数据库数据表打开后
   
    if err.number<>0 then
    response.write "<p align='center'>数据库中暂时无数据</p>"
    end if    
      if rs.eof And rs.bof then
           Response.Write "<p align='center'>对不起，没有符合条件记录！</p>"
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
'=============分页类代码结束%>

 <%'=============循环体开始
   sub showContent
   dim i
   i=0  %>
<table width='100%' height="22" border=0 align=center cellpadding=0 cellspacing=1 bgcolor="#6699cc">
<tr bgcolor="#6699cc">
  <td height="28" align="center"><div align="left">
    <table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr><td height="28" colspan="2"><div align="center" ><font color="#FFFFFF"><b>广告列表</b></font></div></td>
        
        <td>
              <div align="right">
               
              </div></td>
      </tr>
    </table>
  </div>        </td>
  </tr>
</table>          <TR>
          <TD width="60%">&nbsp;&nbsp;广告所属城市：
			  <SELECT name="city" onChange="location=this.value">
				  <OPTION value="?all=ok">全部
				  </OPTION>
				  <OPTION value="?oneid=0&twoid=0&threeid=0" <%if oneid=0 Then%>selected<%End if%>>总站</OPTION>
				  <%Set Rsc=Conn.Execute("Select id,twoid,threeid,city from DNJY_city order by id,twoid,threeid")
				  Do While Not Rsc.Eof
					  If Rsc(0)<>0 and Rsc(1)=0 and Rsc(2)=0 Then%>
					  <OPTION style="background-color:#F8F4F5 ;color: #FF0000" value="?oneid=<%=Rsc(0)%>" <%if Rsc(0)=oneid and 0=twoid Then%>selected<%End if%>><%=Rsc(3)%></OPTION>
					  <%end if
					  If Rsc(0)<>0 and Rsc(1)<>0 and Rsc(2)=0 Then%>
					  <OPTION style="background-color:#F8F4F5 ;color: #0000FF" value="?oneid=<%=Rsc(0)%>&twoid=<%=Rsc(1)%>" <%if Rsc(0)=oneid and Rsc(1)=twoid and 0=threeid Then%>selected<%End if%>>├<%=Rsc(3)%></OPTION>
					  <%end if
					  If Rsc(0)<>0 and Rsc(1)<>0 and Rsc(2)<>0 Then%>
					  <OPTION value="?oneid=<%=Rsc(0)%>&twoid=<%=Rsc(1)%>&threeid=<%=Rsc(2)%>" <%if Rsc(0)=oneid and Rsc(1)=twoid and Rsc(2)=threeid Then%>selected<%End if%>>├├<%=Rsc(3)%></OPTION>
					  <%End if
					  Rsc.MoveNext
				  Loop
				  %>
              </SELECT>
			   </TD><TD  width="40%" height="20" align="center"> &nbsp;&nbsp;&nbsp;&nbsp;<b>一键刷新生成广告JS代码：</b>&nbsp;&nbsp;&nbsp;<a href="createjs_all.asp?city_oneid=<%=rs("city_oneid")%>&city_twoid=<%=rs("city_twoid")%>&city_threeid=<%=rs("city_threeid")%>&city=ok">按地区进行</a> &nbsp;&nbsp;&nbsp;&nbsp; <a href="createjs_all.asp">按全站进行</a></TD>
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
		<td width="10%">选择</td>
          <td width="10%" height="25"><font color="#0000FF">ID号：<%=ADID%></font> <%if IsStop(ADViews,ADStopViews,ADStopHits,ADHits,ADStopDate) Then Response.Write(" <font color=#FF0000>(已过期)") End If%></td>
          <td width="10%">显示：<%=ADViews%></td>
          <td width="10%">点击：<%=ADHits%></td>
		  <td width="15%"><b>活动：</b><%if adkg=1 then%>√<%else%><font color="#FF0000">ⅹ</font><%end if%>&nbsp;&nbsp;&nbsp;&nbsp;<%if adkg=1 then%><a href="createjs.asp?ad=zt&ADID=<%=ADID%>&city_oneid=<%=rs("city_oneid")%>&city_twoid=<%=rs("city_twoid")%>&city_threeid=<%=rs("city_threeid")%>&url=list.asp&page=<%=request("page")%>">暂停</a><%else%><a href="createjs.asp?ad=hd&ADID=<%=ADID%>&city_oneid=<%=rs("city_oneid")%>&city_twoid=<%=rs("city_twoid")%>&city_threeid=<%=rs("city_threeid")%>&url=list.asp&page=<%=request("page")%>"><font color="#FF0000">活动</font></a><%end if%>&nbsp;&nbsp
          <td width="40%"><b>管理：</b><a href="edit.asp?id=<%=id%>">编辑</a>&nbsp;&nbsp;<a href="createjs.asp?ADID=<%=ADID%>&city_oneid=<%=rs("city_oneid")%>&city_twoid=<%=rs("city_twoid")%>&city_threeid=<%=rs("city_threeid")%>">刷新</a>&nbsp;&nbsp;<a href="../ad_info.asp?ADID=<%=ADID%>&city_oneid=<%=rs("city_oneid")%>&city_twoid=<%=rs("city_twoid")%>&city_threeid=<%=rs("city_threeid")%>" target="_blank">预览</a>&nbsp;&nbsp;<span id="followImg<%=i%>" style="CURSOR: hand" onClick="loadThreadFollow(<%=i%>,5)">代码</span>&nbsp;&nbsp;<span id="followImg<%=i%>" style="CURSOR: hand" onClick="loadThreadFollow(<%=i%>,5)">联系广告主</span>
          <%If IsSys = 1 Then
			response.Write "<font color=#FF0000>默认广告不能删除</font>"
		Else
		response.Write "<a href='del.asp?id="&id&"' onclick='return delad();'>删除</a>"
		end if%>
          </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="50%" height="25" colspan=2>广告类型：<%=ShowAdType(ADType,rs("ADSrc"))%></td><td height="25" colspan=2>广告规格：<%=ADWidth%>×<%=ADHeight%></td>
          <td width="50%" colspan=2>显示地址：
          <%If ADType<>6 Then%><a href=../<%=ADSrc%> target=_blank><font color="#000000"><%=ADSrc%></a><%else%>不显示<%end if%></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td height="25" colspan=4>链接地址：<a href=<%=ADLink%> target=_blank><font color="#000000"><%=ADLink%></a></td>
          <td colspan=2>到期日期：<%If now()>=ADStopDate then%><font color="#FF0000"><%=ADStopDate%></font><%else%><%=ADStopDate%><%End if%>&nbsp;&nbsp;&nbsp;&nbsp;提示文字：<%=ADAlt%></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td height="25" colspan=4>所属地区：<%If rs("city_oneid")=0 Then%>总站<%else%><%=Conn.Execute("select city from DNJY_city where twoid=0 and id="&rs("city_oneid"))(0)%><%If rs("city_twoid")>0 then%>/<%=Conn.Execute("select city from DNJY_city where id="&rs("city_oneid")&" and threeid=0 and twoid="&rs("city_twoid")&"")(0)%><%End If%><% If rs("city_threeid")>0 then%>/<%=Conn.Execute("select city from DNJY_city where id="&rs("city_oneid")&" and twoid="&rs("city_twoid")&" and threeid="&rs("city_threeid"))(0)%><%End if%><%End if%></td>
          <td colspan=2>备    注：<%=ADNote%></td>                         
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
  <tr><td height="28" colspan="2"><div align="center" ><b>广告[<%=ADID%>]代码</b></div></td> <td height="28"><div align="center" ><b><b>本广告广告主联系信息</b></b></div></td> 
  </tr>
  <TR>
	<TD width="250"><textarea name="S1" cols="50" rows="6" class="input2"><!--JS固定地区调用方式（无JS文件无法显示，动静态页面均可用）代码开始-->
  <script src="ads/JS/<%=ADID%>_<%=rs("city_oneid")%>_<%=rs("city_twoid")%>_<%=rs("city_threeid")%>.js"></script>
  <!--广告代码结束-->
      </textarea>
        <textarea name="S1" cols="50" rows="6" class="input2"><!--JS动态地区调用方式（无JS文件用总站广告代替，只能动态页面用）代码开始-->
  <script src="&lt;%=adjs_path("ads/js","<%=ADID%>",c1&"_"&c2&"_"&c3)%>"></script>
  <!--广告代码结束-->
      </textarea> </TD>
	<TD width="250">      <textarea name="textarea" cols="50" rows="6" class="input2"><!--ASP固定地区调用方式代码（只能动态页面用）开始-->
  <script src="ads/ad.asp?adid=<%=ADID%>&city_oneid=<%=rs("city_oneid")%>&city_twoid=<%=rs("city_twoid")%>&city_threeid=<%=rs("city_threeid")%>"></script>
<!--广告代码结束-->
      </textarea>
      <textarea name="textarea" cols="50" rows="6" class="input2"><!--ASP动态地区调用方式代码（只能动态页面用）开始-->
  <script src="ads/ad.asp?adid=<%=ADID%>&city_oneid=&lt;%=c1%>&city_twoid=&lt;%=c2%>&city_threeid=&lt;%=c3%>"></script>
<!--广告代码结束-->
      </textarea></TD>
	<TD width="250">
	姓名:<%=ADhost1%><br>电话:<%=ADhost2%><br>手机:<%=ADhost3%><br>Q &nbsp;Q:<%=ADhost4%><br>邮箱:<%=ADhost5%><br>地址:<%=ADhost6%><br>邮编:<%=ADhost7%><br>广告费:<%=ADhost8%> 元<br>备注:<%=ADhost9%></TD>
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
<%'=============放置分页显示开始 
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
    首 页 上一页
    <% Else  %>
    <a href=<% = filename %>?page=1&all=<%=Request("all")%>&<%=Temp%>>首 页</a>
    <a href=<% = filename %>?page=<% = CurrentPage-1 %>&all=<%=Request("all")%>&<%=Temp%>>上一页</a> 
    <% End If 
    If n-currentpage<1 Then  %>
    下一页 尾 页
    <%  Else  %>
    <a href=<% = filename %>?page=<% = (CurrentPage+1) %>&all=<%=Request("all")%>&<%=Temp%>>下一页</a> 
    <a href=<% = filename %>?page=<% = n %>&all=<%=Request("all")%>&<%=Temp%>>尾 页</a>&nbsp;&nbsp;
    <% End If  %>
 页次：<b><font color=red><% = CurrentPage %></font></b>/<b><% = n %></b>页 <b><%=maxperpage%></b>个记录/页  共<b><%=totalnumber %></b>个记录    
转到：<select name="cndok" onchange="javascript:location=this.options[this.selectedIndex].value;">
<%
for i = 1 to n
if i = CurrentPage then%>
<option value="<% = filename %>?page=<%=i%>&all=<%=Request("all")%>&<%=Temp%>" selected>第<%=i%>页</option>  
<%else%>
<option value="<% = filename %>?page=<%=i%>&all=<%=Request("all")%>&<%=Temp%>">第<%=i%>页</option>  
<%
end if
next
%>
</select></font>
</form>
<%End Function 
'=============放置分页显示结束%>
</body>
</html>

<%'检测是否过期
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

'判断广告类型
function ShowAdType(ADType,ADSrc)
	Dim ADExt
	ADExt="图片"
	If InStr(1,ADSrc,".swf",1)>0 Then ADExt="FLASH"
	Select Case ADType
		Case 1
			ShowAdType="普通"&ADExt
		Case 2
			ShowAdType="全屏浮动"&ADExt
		Case 3
			ShowAdType="上下浮动 - 右"&ADExt
		Case 4
			ShowAdType="上下浮动 - 左"&ADExt
		Case 5
			ShowAdType="渐隐消失"&ADExt
		Case 6
			ShowAdType="网页对话框"
		Case 7
			ShowAdType="移动透明对话框"
		Case 8
			ShowAdType="打开新窗口"
		Case 9
			ShowAdType="弹出新窗口"
		Case 10 
			ShowAdType="对联式广告（固定）"
        Case 11 
			ShowAdType="对联式广告（浮动）"	
        Case 12 
			ShowAdType="走马灯文字广告"
        Case 13 
			ShowAdType="普通文字广告"
        Case 14 
			ShowAdType="自写代码广告"
        Case 15 
			ShowAdType="Flash图片广告"
        Case 16 
			ShowAdType="视频广告"				
		Case else
			ShowAdType="<font color=red><b>错误!将不能正确显示</b>"
	End Select
end function
%>