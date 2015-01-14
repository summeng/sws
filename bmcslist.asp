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



</head>
<body topmargin="0" leftmargin="0" style="text-align:center;">
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0" class="tablest">

  <tr>
    <td width="200" height="100%" valign="top" >
<table width="200" border="0" align="left" cellPadding="0" cellSpacing="0" bordercolor="#111111" class="font_10_e_blue" style="border-collapse: collapse">

         <tr>
           <td vAlign="top"><div align="center"> 
             
  <!---分页左侧下广告开始---->  
  <table align="center" bgcolor="#ffffff" >
  <tr>  
   <td ><script language=javascript src="<%=adjs_path("ads/js","info5",c1&"_"&c2&"_"&c3)%>"></script></td>
  </tr>
  </table>   
  </div></td>
         </tr>         
       </table>
	   </td>
 
<td width="760" valign="top" bgcolor="#FFFFFF">
	<table width="760" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF"> 
	<!---上部通栏广告开始---->
  <table align="center" bgcolor="#ffffff">
  <tr>  
   <td ><script language=javascript src="<%=adjs_path("ads/js","fygg",c1&"_"&c2&"_"&c3)%>"></script></td>
  </tr>
  </table>
 
  <!---上部通栏广告结束----> 

<table width="760" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF"> 
<tr><TD>
<%=bmfw(c1,c2,c3,0,6,100,20,12,6,1,1,1,0)%>
</TD></tr>
</table>
</td>
<%Call CloseDB(conn)%>
  <table width="980" cellpadding="0" cellspacing="0">
     <tr>
       <td><!--#include file=footer.asp--></td>
     </tr>
</table>
</body>
</html>