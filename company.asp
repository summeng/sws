<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=top.asp-->

<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Call OpenConn
function HTMLEncode(fString)
if not isnull(fString) then
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")
    fString = Replace(fString, CHR(32), "&nbsp;")
    fString = Replace(fString, CHR(9), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR> ")
    HTMLEncode = fString
end if
end function
Function GetIPAddressData(IP)
	dim sip,num,str1,str2,str3,str4,Rs
	if ip<>"" then
		sip=ip
		If inStr(sip,".") = 0 Then Exit Function
		str1=left(sip,instr(sip,".")-1)
		sip=mid(sip,instr(sip,".")+1)
		If inStr(sip,".") = 0 Then Exit Function
		str2=left(sip,instr(sip,".")-1)
		sip=mid(sip,instr(sip,".")+1)
		If inStr(sip,".") = 0 Then Exit Function
		str3=left(sip,instr(sip,".")-1)
		str4=mid(sip,instr(sip,".")+1)
		if isNumeric(str1)=0 or isNumeric(str2)=0 or isNumeric(str3)=0 or isNumeric(str4)=0 then
		else
		num=cint(str1)*256*256*256+cint(str2)*256*256+cint(str3)*256+cint(str4)-1
		end if
	else
		ip="0.0.0.0"
		num=0
		str1="0.0"
	End If
	'If abc = 0 or abc<1 Then Exit Function
	Set Rs = Conn.ExeCute("select top 1 country,city from [ip] where ip1 <=" & num & " and ip2 >=" & num)
	If Not Rs.Eof Then
		GetIPAddressData = Rs("Country") & rs("city")
	end If
	Rs.Close
	Set Rs = Nothing
	If GetIPAddressData = "" Then GetIPAddressData="δ&#1450;"
end function
%>
<%
dim rsqy,k,vip,mpkg,mpname,qynavigate
username=trim(request("username"))
set rsqy = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_user] where mpkg=1 and username='"&username&"'" 
rsqy.open sql,conn,1,1
if rsqy.eof then
response.write "<script language=JavaScript>" & chr(13) & "alert('你查看的用户没有建立名片或者名片还没有开放哦！');" &"window.location='index.asp'" & "</script>"
response.end
end If
vip=rsqy("vip")
%>

<style>
<!--
.LoginInput {
	BORDER-RIGHT: #eeeeee 1px solid; BORDER-TOP: #999999 1px solid; FONT: 12px "宋体", "Verdana", "Arial"; BORDER-LEFT: #999999 1px solid; BORDER-BOTTOM: #eeeeee 1px solid; HEIGHT: 17px; BACKGROUND-COLOR: #f6f6f6
}
-->
</style>
</head>
<BODY topmargin="0" leftmargin="0" bgColor=#303880>
<CENTER>
<table width="980" border="0" align="center" cellpadding="0" cellspacing="1" class="dq1">
  <tr>
    <td ><script src="<%=adjs_path("ads/js","toppic",c1&"_"&c2&"_"&c3)%>"></script></td>
  </tr>
</table>
<table  border="0" cellspacing="0" cellpadding="0">
    <tr>
     <td  height="7"></td>
    </tr>
</table>
<table style="margin-bottom:8px;" width="980" border="0" align="center" cellpadding="0" cellspacing="1" class="dq1">
  <tr>
    <td height="25" class="dq4" style="padding-left:6px;"><b>行 业 分 类</b> &nbsp;&nbsp;&nbsp;&nbsp;<a href="qyclass.asp?<%=C_ID%>&t=0,0,0">全部分类</a></td>
  </tr>
												<tr>
													<td colspan="4"><%=m_more(1,0,1,1,0,0,1,0,3,11,9,10,8)%></td>
												</tr>
												
</table>


<div align="center">
	<table border="0" style="border-collapse: collapse" width="980" cellpadding="0" height="4" bgcolor="#FFFFFF">
		<tr>
			<td></td>
		</tr>
	</table>
</div>
<div align="center">
<table border="0" width="980" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
	<tr>
		<td width="210" valign="top">
      <TABLE cellSpacing=0 cellPadding=0 border=0 height="360" class="dq1">
        <TBODY>
        <TR>
          <TD align=middle class="dq4" width=210 height=30><b>最 新 加 盟</b></TD></TR>
        <TR>
          <TD valign="top">
<%=dq_more(0,0,0,1,15,3,1,1,13,10,18,0,0,0)%></TD></TR>
        </TBODY></TABLE>
      <table width="210" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
      <TABLE cellSpacing=0 cellPadding=0 border=0 height="360" class="dq1">
        <TBODY>
        <TR>
          <TD align=middle class="dq4" width=210 height=30><b>点 击 热 门</b></TD></TR>
        <TR>
          <TD valign="top">
<%=dq_more(0,0,0,1,15,2,0,2,11,10,18,1,0,0)%></TD></TR>
        </TBODY></TABLE>
		</td>
		<td width="5" valign="top">　</td>
		<td width="75%" valign="top">
      <TABLE borderColor=#0D81BC cellSpacing=0 cellPadding=0 border=0 width="760">
        <TBODY>

        <TR>
          <TD align=middle><!--开始-->
		  <table cellSpacing="0" cellPadding="0" width="760" border="0"  class="dq1">
			<tr>				
				<td align="middle" width="198"  height="30" class="dq4"><b>企 业 名 片</b></td>				
			</tr>
			<tr vAlign="top" align="middle">
				<td colSpan="5">
				<table cellSpacing="0" cellPadding="0" width="100%" border="0">
					<tr>
						<td>
						<table cellSpacing="0" cellPadding="0" width="100%" border="0">
							<tr>
								<td><!--开始--><TABLE cellSpacing=0 cellPadding=0 width="72%" align=center border=0 height="156">
        <TBODY>
        <TR>
          <TD height=10></TD></TR>
        <TR>
          <TD vAlign=top width="90%" height=146>
      <div align="center">

<%Dim oColor
If rsqy("mpfg")=1 Then oColor="A8D7E4"
If rsqy("mpfg")=2 Then oColor="6996FF"
If rsqy("mpfg")=3 Then oColor="23C623"
If rsqy("mpfg")=4 Then oColor="FFA643"
If rsqy("mpfg")=5 Then oColor="FF9FD4"
%>
      <TABLE cellSpacing=0 cellPadding=0 width="510" border=0 id="table10">
        <TBODY>
        <TR>
          <TD width="37">
			<img border="0" src="images/qyimg/qyls/mpa<%=rsqy("mpfg")%>.gif" width="37" height="35"></TD>
          <TD width="650" background="images/qyimg/qyls/mpb<%=rsqy("mpfg")%>.gif" valign="bottom">
			<DIV style="FILTER: dropshadow(color=#000000, offx=1, offy=1, positive=2); WIDTH: 650px; COLOR:#ffffff; FONT-FAMILY: Arial; POSITION: relative; height:27px">
			<STRONG><p align=center style="line-height: 250%"><FONT style="FONT-SIZE: 26px" face="隶书" color="#FFFFFF"><%=rsqy("mpname")%></FONT></p></STRONG></DIV></TD>
          <TD width="10">
			<img border="0" src="images/qyimg/qyls/mpc<%=rsqy("mpfg")%>.gif" width="10" height="35"></TD></TR>
        <TR>
          <TD bgColor=#<%=oColor%> colspan="3">
            <div align="center">
			<TABLE cellPadding=4 width="650" border=0 id="table11"><tr><td > <%if rsqy("logo")<>"" then %><img   border="0" src="<%=rsqy("logo")%>" title="商家标志" height="194" width="240"><%else%><img src="images/qyimg/nopic.jpg" align="middle" border=0 alt="没有店标" height="194" width="240"><%end if%></td><td><b>单位简介：</b><%=rsqy("mpjj")%></td></tr></table>
            <TABLE cellPadding=4 width="650" border=0 id="table11">
              <TBODY>
              <TR>
                <TD><b>所属行业：</b><%belongtype="<a href=""qyclass.asp?t="&rsqy("type_oneid")&",0,0"">"&rsqy("type_one")&"</a>"
	IF rsqy("type_two")<>"" and not isnull(rsqy("type_two")) Then belongtype=belongtype&"/<a href=""qyclass.asp?t="&rsqy("type_oneid")&","&rsqy("type_twoid")&",0"">"&rsqy("type_two")&"</a>"
	IF rsqy("type_three")<>"" and not isnull(rsqy("type_three")) Then belongtype=belongtype&"/<a href=""qyclass.asp?t="&rsqy("type_oneid")&","&rsqy("type_twoid")&","&rsqy("type_threeid")&""">"&rsqy("type_three")&"</a>"
	response.write belongtype%></TD><TD><b>所在地区：</b><%qynavigate="<a href=""qyclass.asp?c="&rsqy("city_oneid")&",0,0"">"&rsqy("city_one")&"</a>"
	IF rsqy("city_two")<>"" and not isnull(rsqy("city_two")) Then qynavigate=qynavigate&"/<a href=""qyclass.asp?c="&rsqy("city_oneid")&","&rsqy("city_twoid")&",0"">"&rsqy("city_two")&"</a>"
	IF rsqy("city_three")<>"" and not isnull(rsqy("city_three")) Then qynavigate=qynavigate&"/<a href=""qyclass.asp?c="&rsqy("city_oneid")&","&rsqy("city_twoid")&","&rsqy("city_threeid")&""">"&rsqy("city_three")&"</a>"
	response.write qynavigate%></TD></TR>
             <TR>
                <TD width="254"><b>主&nbsp;&nbsp;&nbsp;&nbsp;营：</b><%=rsqy("zhuying")%></TD>
                <TD><b>公司网站：</b><a href="http://<%=rsqy("http")%>" target="_blank"><%=rsqy("http")%></a></TD></TR>
              <TR>
                <TD width="254"><b>联 系 人：</b><%=rsqy("name")%></TD>
                <TD><b>看其网店：</b><%if rsqy("sdname")<>"" And rsqy("sdkg")=1 then%><a  href="user/com.asp?com=<%=rsqy("username")%>&<%=C_ID%>" target="_blank"><img src="img/dian.gif" title="已开有网店" border="0"></a><%else%>此会员未开有网店<%end if%></TD></TR>
				<TR>
                <TD width="254"><b>邮政编码：</b><%=rsqy("youbian")%></TD>
                <TD><b>看其信息：</b><a href="preview.asp?username=<%=rsqy("username")%>" target="_blank">查看其发布的信息</a></TD></TR>
              <TR>
				<TR>
                <TD width="254"><b>通讯地址：</b><%=rsqy("dizhi")%></TD>
                <TD><b>站内短信：</b><a href="messh.asp?name=<%=rsqy("username")%>" target="_blank">发站内短信给他</a></TD></TR>
              <TR>
				<TR>
                <TD width="254"><b>联系固话：</b><%=rsqy("dianhua")%></TD>
                <TD><b>电子信箱：</b><%if rsqy("email")<> "" and rsqy("email")<> "@" then%>
                        <INPUT name="button" TYPE=button class="inputb" style="color: #666666; background-color: #E1E2DC" ONCLICK="window.open('user/user_mail.asp?username=<%=rsqy("username")%>&email=<%=rsqy("email")%>&mailzt=<%=rsqy("mpname")%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=650,height=300,left=300,top=100')" VALUE="发邮件交流"></a>
                        <%end if%></TD></TR>
              <TR>
              <TR>
				<TR>
                <TD width="254"><b>联系手机：</b><%=rsqy("CompPhone")%></TD>
                <TD><b>诚信记录：</b><%=rsqy("goodfaith")%> 次 <img src="img/hp.jpg"></TD></TR>
              <TR>
				<TR>
                <TD width="254"><b>传&nbsp;&nbsp;&nbsp;真：</b><%=rsqy("fax")%></TD>
                <TD><b>失信记录：</b><%=rsqy("bpromises")%> 次 <img src="img/cp.jpg"></TD></TR>
              <TR>
              <TR>
                <TD width="254"><b>联络QQ号：</b><%=rsqy("qq")%></font><%if rsqy("qq")<>"" then %>&nbsp; <a target=blank href=http://wpa.qq.com/msgrd?V=1&Uin=<%=rsqy("qq")%>&Menu=yes><img border="0" SRC=http://wpa.qq.com/pa?p=1:<%=rsqy("qq")%>:7 alt="用QQ给他发信息"></a><%end if%></TD>
                <TD><b>网友打分：</b><%=rsqy("wevaluation")%> 分&nbsp;&nbsp;&nbsp;<INPUT TYPE=button VALUE="我要打分" ONCLICK="window.open('user/evaluation.asp?username=<%=rsqy("username")%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbarsper=yes,scrollbars=yes,resizable=no,copyhistory=no,width=400,height=300,left=300,top=100')" style="color: #0060C5; background-color: #E1E2DC">&nbsp;&nbsp;&nbsp;<INPUT TYPE=button VALUE="我要收藏" ONCLICK="window.open('user/collection_shops.asp?scid=<%=rsqy("username")%>&name=<%=rsqy("name")%>&sdname=<%=rsqy("sdname")%>&mpname=<%=rsqy("mpname")%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbarsper=yes,scrollbars=yes,resizable=no,copyhistory=no,width=400,height=300,left=300,top=100')" style="color: #0060C5; background-color: #E1E2DC"></TD></TR>

              <TR>
                <TD valign="middle" colspan="2" style="TABLE-LAYOUT: fixed;word-wrap:break-word;">
				<b>公 告 板：</b><marquee onMouseOver="this.stop()" onMouseOut="this.start()" scrollamount="3" scrolldelay="121colspan="2"" behavior="scroll"   style="font-size: 10pt; " width="600"><%if IsNull(rsqy("comgg")) then%>欢迎光临<%=rsqy("mpname")%>主页<%else%><%=rsqy("comgg")%><%End if%></marquee></TD>
                </TR>

              </TBODY></TABLE></div>
			</TD></TR>
        <TR>
          <TD colspan="3">
			<IMG height=9 alt="" 
            src="images/qyimg/qyls/mpd<%=rsqy("mpfg")%>.gif" 
        width=700></TD></TR></TBODY></TABLE>

　</div>
			</TD>
          </TR>
        </TBODY></TABLE>
<!--结束-->
								</td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
<!--结束--></TD></TR>

        <TR>
          <TD align=middle height=13>
<div class="toleft">
<table width="760" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
<table  id="Table5" cellSpacing="0" cellPadding="5"  border="0" width="100%"  class="dq1">
						<tr>
							<td>
		<table width="100%" border="0" cellspacing="0" cellpadding="0" >
        <tr> 
          <td class="middle" align="center" valign="top"> 
            <div class="top"><span style="float:left;padding-top:2px;"><%=dq_search(1,1,1)%></div>
        </td>
        </tr>
      </table>

							</td>
						</tr>
					</table>
</div>	</TD></TR>
        </TBODY></TABLE>
      <table width="760" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
<!---730广告开始----><table width="100%" border="0" cellspacing="0" cellpadding="0" class="dq1">
        <tr> 
          <td class="middle" align="center" valign="top"> 
            <div class="top"> 
<script src="<%=adjs_path("ads/js","syzb730",c1&"_"&c2&"_"&c3)%>"></script>
  </div>
        </td>
        </tr>
</table><!---730结束---->
		</td>
	</tr>
	</table>
	</div>
<TABLE cellSpacing=0 cellPadding=0 width=980 align=center border=0>
  <TBODY>
  <TR><%conn.execute "UPDATE DNJY_user SET mphits=mphits+1 WHERE username='"&username&"'" '点击计数
   rsqy.close
   set rsqy=Nothing
   Call CloseDB(conn)%>
    <TD width="769">
    <p align="center"><!--#include file=footer.asp--></TD></TR>                         
  <TR>
<TD width="769"></TD></TR></TBODY></TABLE>
</body>
</html>
