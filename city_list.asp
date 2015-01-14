<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file="top.asp"-->
<TABLE width="980"  border="0" align="center" cellpadding="0" cellspacing="0">
  <TR>
    <TD><IMG src="img/city_type.jpg" width="980" height="35"></TD>
  </TR>
</TABLE>
<table width="980"  border="0" align="center" cellpadding="0" cellspacing="2" bgcolor="CACD1A">

  <TR>
    <TD bgcolor="#FFFFFF">
<!--<TABLE width="980">
<TR>
	<TD width="50%" align="center"><b>竞价信息</b><br><%=f_info(2,0,0,0,c1,c2,c3,1,1,1,1,3,1,0,1,1,0,38,10,1,13,11)%></TD>
	<TD width="50%"><img src="img/mapcn1.gif" width="368" height="327" border="0" usemap="#Map" /><img src="img/mapcn2.gif" width="100" height="153" border="0" /></TD>
</TR>
</TABLE>-->
<%=c_more(0,1,1,0,0,0,0,1,0,3,12,11,9,40,20,15)%>
		   <%If sys="" Then sys=0%>
            </span></td>
      </tr>
    </TABLE>

</TABLE>
      <TABLE  border="0" cellpadding="0" cellspacing="0" align="center" id="table17">
        <TR>
          <TD height="8"></TD>
        </TR>
      </TABLE>
      <TITLE></TITLE>
<%Dim webym
webym=Mid(""&web&"",5)%>
<map name="Map" id="Map">
<area shape="rect" coords="31,91,132,111" href="http://xj.<%=webym%>" target="_blank" alt="新疆" />
<area shape="rect" coords="317,40,356,59" href="http://hl.<%=webym%>" target="_blank" alt="黑龙江" />
<area shape="rect" coords="315,73,348,89" href="http://jl.<%=webym%>" target="_blank" alt="吉林" />
<area shape="rect" coords="298,95,328,111" href="http://ln.<%=webym%>" target="_blank" alt="辽宁" />
<area shape="rect" coords="261,102,289,118" href="http://he.<%=webym%>" target="_blank" alt="河北" />
<area shape="rect" coords="151,254,183,270" href="http://yn.<%=webym%>" target="_blank" alt="云南" />
<area shape="rect" coords="56,186,120,205" href="http://xz.<%=webym%>" target="_blank" alt="西藏" />
<area shape="rect" coords="123,148,158,166" href="http://qh.<%=webym%>" target="_blank" alt="青海" />
<area shape="rect" coords="160,130,193,147" href="http://gs.<%=webym%>" target="_blank" alt="甘肃" />
<area shape="rect" coords="167,111,242,126" href="http://nm.<%=webym%>" target="_blank" alt="内蒙古" />
<area shape="rect" coords="243,122,269,136" href="http://bj.<%=webym%>/" target="_blank" alt="北京" />
<area shape="rect" coords="233,138,260,153" href="http://sx.<%=webym%>" target="_blank" alt="山西" />
<area shape="rect" coords="197,141,223,157" href="http://nx.<%=webym%>" target="_blank" alt="宁夏" />
<area shape="rect" coords="212,169,239,185" href="http://sn.<%=webym%>" target="_blank" alt="陕西" />
<area shape="rect" coords="170,202,195,220" href="http://sc.<%=webym%>" target="_blank" alt="四川" />
<area shape="rect" coords="198,199,226,217" href="http://cq.<%=webym%>/" target="_blank" alt="重庆" />
<area shape="rect" coords="196,235,222,250" href="http://gz.<%=webym%>" target="_blank" alt="贵州" />
<area shape="rect" coords="212,263,244,278" href="http://gx.<%=webym%>" target="_blank" alt="广西" />
<area shape="rect" coords="234,222,262,237" href="http://hn.<%=webym%>" target="_blank" alt="湖南" />
<area shape="rect" coords="235,193,260,209" href="http://hb.<%=webym%>" target="_blank" alt="湖北" />
<area shape="rect" coords="245,167,272,182" href="http://ha.<%=webym%>" target="_blank" alt="河南" />
<area shape="rect" coords="273,181,298,196" href="http://ah.<%=webym%>" target="_blank" alt="安徽" />
<area shape="rect" coords="297,166,326,183" href="http://js.<%=webym%>" target="_blank" alt="江苏" />
<area shape="rect" coords="271,146,303,161" href="http://sd.<%=webym%>" target="_blank" alt="山东" />
<area shape="rect" coords="289,126,318,143" href="http://tj.<%=webym%>/" target="_blank" alt="天津" />
<area shape="rect" coords="297,206,323,222" href="http://zj.<%=webym%>" target="_blank" alt="浙江" />
<area shape="rect" coords="323,182,350,197" href="http://sh.<%=webym%>/" target="_blank" alt="上海" />
<area shape="rect" coords="250,254,279,270" href="http://gd.<%=webym%>" target="_blank" alt="广东" />
<area shape="rect" coords="281,235,310,249" href="http://fj.<%=webym%>" target="_blank" alt="福建" />
<area shape="rect" coords="264,216,291,230" href="http://jx.<%=webym%>" target="_blank" alt="江西" />
<area shape="rect" coords="327,253,354,268" href="http://tw.<%=webym%>/" target="_blank" alt="台湾" />
<area shape="rect" coords="275,278,303,292" href="http://hk.<%=webym%>/" target="_blank" alt="香港" />
<area shape="rect" coords="249,289,276,303" href="http://mo.<%=webym%>/" target="_blank" alt="澳门" />
<area shape="rect" coords="245,305,273,319" href="http://hi.<%=webym%>" target="_blank" alt="海南" />
</map>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="footer.asp"-->
