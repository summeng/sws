<!--#include file="sdtop.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Dim m%>
<title><%=rs("sdname")%></title>
<meta name="keywords" content="<%=rs("sdname")%>" />
<meta name="description" content="<%=CutStr(rs("sdjj"),200,"...")%>" />
<table width="960" align="center" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="190" align="left" valign="top" class="top_m_txt01">
<!--#include file="sdleft.asp"-->
</td>
<td width="5"></td>
<td width="765" valign="top">
<table width="765" border="0" cellpadding="0" cellspacing="0" id="line">
      <tr>
        <td valign="top">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td><div id="lt2"><span style="margin-left:50px;">公司简介</span></div></td>
          </tr>
        </table>
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td><div class="aboutpic"><%if rs("logo")<>"" then %> <a target="_blank" title="点击放大&gt;&gt;&gt;" href="../<%=rs("logo")%>"><img border="0" src="../<%=rs("logo")%>" title="点击放大" width="220"height="180"></a><%else%><img src="../images/nopic.gif" align="middle" border=0 alt="没有店标"><%end if%></div><div style="margin:8px; line-height:23px;;font-size:10pt;"><%=HTMLDecode(rs("sdjj"))%></div></td>
            </tr>
          </table></td>
      </tr>
    </table>
<table width="765" border="0" cellpadding="0" cellspacing="0" id="line">
      <tr>
        <td valign="top">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td><div id="lt2"><span style="margin-left:50px;">公司地图</span></div></td>
          </tr>
        </table>
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td><IFRAME border=0 src="Map.asp?com=<%=com%>" frameBorder=0 width=760 scrolling=no height=410></IFRAME></td>
            </tr>
          </table></td>
      </tr>
    </table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td height="5"></td>
            </tr>
          </table>
	<%If jf<adjf Or adjf=0 then%><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td align="center"><script language=javascript src="../<%=adjs_path("ads/js","xxfl1",c1&"_"&c2&"_"&c3)%>"></script></td>
            </tr>
          </table><%End if%>		

</td>
</tr>			  
</table>
<!--#include file="sdend.asp"-->