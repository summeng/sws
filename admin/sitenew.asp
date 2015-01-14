<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=../config.asp-->
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("05")
dim javastr1,rsjs,fs,f,action,hs,ls,btw,hg,bgw,iii,i
action=request.querystring("action")
If trim(request("hs"))<>"" Then hs=CInt(trim(request("hs")))
If trim(request("ls"))<>"" Then ls=CInt(trim(request("ls")))
If trim(request("bgw"))<>"" Then btw=CInt(trim(request("btw")))
If trim(request("hg"))<>"" Then hg=CInt(trim(request("hg")))
If trim(request("bgw"))<>"" Then bgw=CInt(trim(request("bgw")))
If action="Production" Then
	javastr1="document.write('"
	Call OpenConn
             set rsjs=server.createobject("adodb.recordset")
			 sql="select * from DNJY_info where yz=1 and fsohtm=1"
			 if jiaoyi=0 Then sql=sql&" and zt="&jiaoyi
             if overdue=0 Then sql=sql&" and dqsj>"&DN_Now&""
             sql=sql&" order by fbsj "&DN_OrderType&""
			 rsjs.open sql,conn,3,1
	'-------下面的句子是输出并控制生成格式--------------------------------
	iii=0
	i=0
	javastr1=javastr1+"<table border=0 cellpadding=0 cellspacing=0  width="&bgw&">"
	do while not rsjs.bof and not rsjs.eof
	i=i+1
	IF iii Mod ls=0 then javastr1=javastr1&  "<tr height="""&hg&""">"
	javastr1=javastr1+"<td width="""&bgw&""">·<a href=""http://"&web&"/"&strInstallDir&"html/"&server.HTMLEncode(trim(rsjs("id")&" "))&".html"" target=""_blank"">"
	if len(trim(rsjs("biaoti")))>btw then
	javastr1=javastr1+""&server.htmlencode(trim(mid(rsjs("biaoti"),1,btw-2)&" "))&"...</a></td>"
	Else
	javastr1=javastr1+""&server.htmlencode(trim(rsjs("biaoti")&" "))&"</a></td>"	
	End if
	iii=iii+1
	IF iii Mod ls=0 then javastr1=javastr1&  "</tr>"
    if i>=hs then exit do 
	rsjs.movenext
	Loop	
	IF iii Mod ls<>0 then
	Do While not iii Mod ls=0
	javastr1=javastr1&  "<td></td>"
	iii=iii+1
	Loop
	javastr1=javastr1&  "</tr>"
	End IF
	javastr1=javastr1&  "</table>"
'-------控制生成格式结束----------------------------------------

			  rsjs.close
			  set rsjs=nothing	  

javastr1=javastr1+"');"&vbcrlf
		Set fs = CreateObject("Scripting.FileSystemObject")
		
		Set f = fs.CreateTextFile(Server.MapPath("/")&"/"&""&strInstallDir&"sitenew/sitenew.js", True)
		f.write(javastr1)
		f.close
		Set f = nothing
		Set fs = Nothing
		Call CloseDB(conn)
		response.Write "<script language=javascript>alert('JS调用最新信息文件生成成功!');location.replace('sitenew.asp');</script>"

End if%>
<link href="../css/style.css" rel="stylesheet" type="text/css">
<BODY>
<style type="text/css">
<!--
.style1 {color: #FFFFFF;
	font-weight: bold;
}
.style2 {color: #FF0000}
-->
</style>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#6699cc">

  <TR>
    <TD height="20" bgcolor="#799AE1" align="center" colspan="13"><FONT color="#FFFFFF" style="font-size:14px">生 成 最 新 调 用 参 数</FONT></TD>
 </TR>
  <FORM name="form" method="post" action="?action=Production" >
  
    <tr> 
      <td height="30" colspan="3" align="center">&nbsp;</td>
    </tr>

                      <tr>
                        <td height="25">调用条数：<input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="hs" size="10" value="10" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"> 条 </td>
                      </tr> 
                     <tr>
                        <td height="25">调用列数：<input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="ls" size="10" value="1" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"> 列 </td>
                      </tr>
                      </tr> 
                     <tr>
                        <td height="25">标题长度：<input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="btw" size="10" value="20" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"> 字符 </td>
                      </tr>
                     <tr>
                        <td height="25">表格行高：<input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="hg" size="10" value="10" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')">  </td>
                      </tr> 	
                     <tr>
                        <td height="25">表格宽度：<input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="bgw" size="10" value="250" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')">  </td>
                      </tr> 						  
    <tr> 
      <td height="30" colspan="3" align="center"> 
        <input type="submit" name="Submit" value="提交" class="input" >
        
        <input type="reset" name="Submit2" value="重置" class="input"> 
 </td>
    </tr>
  </form>
  <tr> 
    <td height="30" colspan="3" align="center">&nbsp;</td>
  </tr>
</table>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%">
  
    <tr bgcolor="#BDBEDE">
      <td height="28" style="border-bottom-style: solid; border-bottom-width: 1"><span class="style1"><span style="font-size: 9pt"><span style="font-size: 9pt"><font size="3"><font size="2">&nbsp;</font></font>首页代码调用说明</span></span></span></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td style="border-top-style: solid; border-top-width: 1"><table width="98%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td><p>生成的JS是sitenew目录下的相关地区js文件，站内或站外可直接使用下面代码调用，所调用的信息是静态的，更新信息后需要手工生成<br>在文件名加上地区ID号可调用地区信息：sitenew<font color='red'>c1</font>_<font color='red'>c2</font>_<font color='red'>c3</font>.js，其中c1、c2、c3分别为三级地区的ID号。条件是分站管理员已生成相应地区的文件。</p>
           
            <p class="style2">
              <textarea name="textarea" cols="88" rows="6"><!--代码开始-->
<SCRIPT language=JavaScript src="http://<%=web%>/<%=strInstallDir%>sitenew/sitenew.js" type=text/JavaScript></SCRIPT>
<!--代码结束-->
              </textarea>
            </p>
            <p>JS的生成需要FSO支持</p></td>
        </tr>
      </table></td>
    </tr>
	<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%">
  
    <tr bgcolor="#BDBEDE">
      <td height="28" style="border-bottom-style: solid; border-bottom-width: 1">
<b>调用预览：</b>&nbsp;&nbsp;<input type=button value=刷新效果 onClick="location.reload()"><p>
<!--代码开始-->
<SCRIPT language=JavaScript src="../sitenew/sitenew.js" type=text/JavaScript></SCRIPT>
<!--代码结束-->
</td>
        </tr>
      </table
></table>

