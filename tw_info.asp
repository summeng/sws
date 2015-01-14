<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file="top.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室版权所有
'官方客服中心 http://www.ip126.com
'技术支持论坛 http://www.ip126.com/bbs
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
 <table width="980" border="0" align="center" cellspacing="0" cellpadding="0" class="tablest">
        <tr> 
          <td class="middle" width="980" height="45" align="center" valign="top"><div class="top"><span style="float:left;padding-top:9px;"><%=f_search(3,2)%></div></td><td width="30" height="22"><%
		  set rslx=server.createobject("adodb.recordset")
sqllx = "select leixing from DNJY_config "
rslx.open sqllx,conn,1,1
leixing=split(rslx("leixing"),"|")
response.write"<select name='leixing' onChange=""var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location=jmpURL;} else {this.selectedIndex=0 ;}"" ><option value='?leixing='''>交易类型</option>"
response.write"<option value='tw_info.asp?leixing=&"&CT_ID&"&xsfs=1'>全部类型</option>"
for x=0 to ubound(leixing)
response.write"<option value='tw_info.asp?leixing="&leixing(x)&"&"&CT_ID&"&xsfs=1'>"&leixing(x)&"</option>"
next
response.write"</select>"
rslx.close
set rslx=Nothing
%></td>
         
        </tr>
      </table>
<table width="980" border="0" cellspacing="1" cellpadding="1" align="center">
<tr><%Dim xxsx
xxsx=Cint(SafeRequest("xxsx",1))%>
	<td><%=tw_info(xxsx,t1,t2,t3,c1,c2,c3,0,1,1,0,0,0,1,1,20,1,13,180,180,200,250,30,5,1)%></td>
</tr>
</table>
<%Call CloseDB(conn)%>
<!--#include file=footer.asp-->
