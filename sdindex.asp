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
<div align="center">
<table style="margin-bottom:8px;" width="980" border="0" align="center" cellpadding="0" cellspacing="1" class="dq1">
  <tr>
    <td height="25" class="dq2" style="padding-left:6px;"> 地区导航：<a href="sdindex.asp?c=0,0,0&<%=T_ID%>">总站-></a></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF" style="padding:8px;"><%=W_vassal_c_t(0,0,c1,c2,c3,t1,t2,t3,0,1,0,10,0,22)%></td>
  </tr>
</table>
</div>
<div align="center">
<table border="0" width="980" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
	<tr>
		<td width="210" valign="top">
      <TABLE cellSpacing=0 cellPadding=0 border=0 class="dq1">
        <TBODY>
        <TR>
          <TD align=middle class="dq4" height=30><b>最 新 加 盟</b></TD></TR>
        <TR>
          <TD width="210" height=200 valign="top">
<%=dq_more(0,0,0,0,10,3,1,1,13,10,18,0,0,0)%>
			</TD></TR>  </table>
      <table width="200" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
      <TABLE cellSpacing=0 cellPadding=0 border=0 class="dq1">
        <TR>
          <TD>
      <TABLE borderColor=#0D81BC cellSpacing=0 cellPadding=0 border=0>
        <TBODY>
        <TR>
          <TD align=middle class="dq4" height=30><b>点 击 热 门</b></TD></TR>
        <TR>
          <TD width="215" height=200 valign="top">
<%=dq_more(0,0,0,0,10,2,0,2,11,10,18,1,0,0)%></TD></TR>
        </TBODY></TABLE>
      	</TD></TR></TBODY></TABLE></td>
		<td width="1%" valign="top">　</td>
		<td width="77%" valign="top">
      <TABLE cellSpacing=0 cellPadding=0 border=0 width="760"  class="dq1">
        <TBODY>
        
        <TR>
          <TD align=middle><!--开始--><table cellSpacing="0" cellPadding="0" width="760" border="0">
			<tr>
				<td align="middle" width="114" height="30" class="dq4">
				<b>店 铺 分 类</b></td>
			</tr>
			<tr vAlign="top" align="middle">
				<td  colSpan="5">
				<table cellSpacing="0" cellPadding="0" width="100%" border="0">
					<tr>
						<td>
						<table cellSpacing="0" cellPadding="0" width="100%" border="0">
							<tr>
								<td><!--开始分类-->
<div align="center">
<TABLE cellPadding=0 width=98% border=0>
              <TBODY>
              <TR align=left>
<%=m_more(0,0,1,1,0,0,1,0,3,11,9,10,8)%></TR>

              </TBODY></TABLE>
</div>
<!--结束分类-->
								</td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
      <table width="760" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
	<table  id="Table5" cellSpacing="0" cellPadding="5"  border="0" width="100%" class="dq1">
						<tr>
							<td>
		<table width="100%" border="0" cellspacing="0" cellpadding="0" >
        <tr> 
          <td class="middle" align="center" valign="top"> 
            <div class="top"><span style="float:left;padding-top:2px;"><%=dq_search(0,1,1)%></div>
        </td>
        </tr>
      </table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" >
        <tr> 
          <td class="middle" align="center" valign="top"> 
            <div class="top"> <!---730广告开始---->
<script src="<%=adjs_path("ads/js","syzb730",c1&"_"&c2&"_"&c3)%>"></script>
 <!---730结束----> </div>
        </td>
        </tr>
</table>
							</td>
						</tr>
					</table>
<!--结束--></TD></TR>
        </TBODY></TABLE>
		</td>
	</tr>
	</table>
	</div><!--开始-->
<div align="center">
	<table border="0" style="border-collapse: collapse" width="980" bgcolor=#ffffff>
		<tr>
			<td align=center></td>
		</tr>
		<tr>
			<td background="images/qyimg/qyls/2005_soft_mainbg.gif" height="7"></td>
		</tr>
	</table>
</div>

<div align="center">
	<table cellSpacing="0" cellPadding="0" width="980" border="0"  class="dq1">
			<tr>				
				<td align="middle" width="198" height="30" class="dq4"><b>分 类 导 航</b></td>				
			</tr>

<tr vAlign="top" align="middle">
<td  colSpan="5"  bgcolor="#EEEEEE">
<TABLE height=184 cellSpacing=0 cellPadding=0 width=980 align=center border=0 >
<TBODY>
<TR>
<TD width=980 style="font-size: 12px; color: #000000; font-family: Verdana, Arial, Helvetica, sans-serif">
<TABLE>
<TBODY>
<TR><%=dq_more_dh(0,4,12,10,1,1,1,16,10,13,0,0,0)%>
</TR></TBODY></TABLE></TD>
<TD align=middle width=2 style="font-size: 12px; color: #000000; font-family: Verdana, Arial, Helvetica, sans-serif"></TD></TR></TBODY></TABLE>
</td>
</tr>
</TABLE>    
</div>

<div align="center">
	<table cellSpacing="0" cellPadding="0" width="980" border="0"  class="dq1">
			<tr>				
				<td align="middle" width="198" height="30" class="dq4"><b>共 享 资 源</b></td>				
			</tr>

<tr vAlign="top" align="middle">
<td  colSpan="5"  bgcolor="#EEEEEE">
<TABLE height=184 cellSpacing=0 cellPadding=0 width=980 align=center border=0 >
<TBODY>
<TR>
<TD width=980 style="font-size: 12px; color: #000000; font-family: Verdana, Arial, Helvetica, sans-serif">
<TABLE>
<TBODY>
<TR><%=vipnews("",0,0,1,1,1,0,12,30,15,3,15,10,1)%>
</TR></TBODY></TABLE></TD>
<TD align=middle width=2 style="font-size: 12px; color: #000000; font-family: Verdana, Arial, Helvetica, sans-serif"></TD></TR></TBODY></TABLE>
</td>
</tr>
</TABLE>    
</div>

<!--结束-->
<TABLE cellSpacing=0 cellPadding=0 width=980 align=center border=0>
  <TBODY>
  <TR>
    <TD width="769">
    <p align="center"><!--#include file=footer.asp--></TD></TR>                         
  <TR>
<TD width="769"></TD></TR></TBODY></TABLE>
</body>
</html>
