<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=top.asp-->
<!--#include file=usercookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Dim isse,ClassName,se_ClassName,selecttype,searchkeys,pp,turl,modname,delname,caozuocs,page,pages,j,submitname,pagecanshu
username=request.cookies("dick")("username")
isse=ReplaceUnsafe(request.QueryString("isse"))
ClassName=ReplaceUnsafe(Request.QueryString("ClassName"))
se_ClassName=ReplaceUnsafe(Request.Form("se_ClassName"))
if se_ClassName="" then se_ClassName=ReplaceUnsafe(Request.QueryString("se_ClassName"))
selecttype=ReplaceUnsafe(Request.Form("selecttype"))
if selecttype="" then selecttype=ReplaceUnsafe(Request.QueryString("selecttype"))
searchkeys=ReplaceUnsafe(Request.Form("searchkeys"))
if searchkeys="" then searchkeys=ReplaceUnsafe(Request.QueryString("searchkeys"))
%>
<html>
<head>
<title>�޸���Ϣ�б�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
<script language="JavaScript" type="text/JavaScript">
function checkform()
{
  if (document.postart.searchkeys.value=="")
  {
    alert("�������ѯ������");
	document.postart.searchkeys.focus();
	return false;
  }

  
}
</script>

<script language = "JavaScript">
function go(URL,cfmText)
{
	var ret;
	ret = confirm(cfmText);
	if(ret!=false)window.location=URL;
}
function CheckAll(form)
	{
	for (var i=0;i<form.elements.length;i++){
	var e = form.elements[i];
	if (e.name != 'chkall')
		e.checked = form.chkall.checked;
		}
	}
 </script>
</head>
<body>

  <table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse">
    <tr>
      <td width="214" height="356" valign="top" bgcolor="#FFFFFF"><div align="center"><!--#include file=userleft.asp--></div></td>
	  <td width="5">&nbsp;</td>
      <td width="760" align="center" valign="top" bgcolor="#FFFFFF">
	<!---�ϲ�ͨ����濪ʼ---->
  <table align="center" bgcolor="#ffffff">
  <tr>  
   <td ><script language=javascript src="<%=adjs_path("ads/js","fygg",c1&"_"&c2&"_"&c3)%>"></script></td>
  </tr>
  </table>
  <!---�ϲ�ͨ��������----> 
  <%if isse=0 then
sql="select *,newsShared from DNJY_userNews where username='"&username&"' and ClassName='"& ClassName &"' order by id desc"
end if

if isse=1 Then
	if selecttype=2 then
		if not IsNumeric(searchkeys) then
			response.write"<SCRIPT language=JavaScript>alert('ID��Ӧ��Ϊ���֣�');"
			response.write"javascript:history.go(-1)</SCRIPT>"
			response.end
		end if
		sql="select * from DNJY_userNews where username='"&username&"' and id="&searchkeys&" order by id"
	end if
	if selecttype=1 then
		if se_classname="" then
			sql="select * from DNJY_userNews where username='"&username&"' and title like '%"&searchkeys&"%' order by ClassName,id"
		else
			sql="select * from DNJY_userNews where username='"&username&"' and ClassName='"& se_ClassName &"' and title like '%"&searchkeys&"%' order by id"
		end if
	end if
end if%>
<table width="760" align="center" bgcolor="#ffffff">
    <tr>
      <td><table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td><form name="postart" method="post" action="UserNews_Modifylist.asp?isse=1" onSubmit="return checkform()">
              <br>
              ������Ϣ��
			  
			<script language = "JavaScript">
			  function typechange()
			{ choiceid=postart.selecttype.selectedIndex
   		 
				if(postart.selecttype.options[choiceid].value=="2"){
					
					postart.se_ClassName.disabled=true;
					document.all("idh").style.display="inline"  
					document.all("gjz").style.display="none"
					}
				if(postart.selecttype.value=="1"){
					
					postart.se_ClassName.disabled=false;
					document.all("idh").style.display="none" 
					document.all("gjz").style.display="inline"
				}
		
			 }
			 </script>
                          <select name="selecttype" onChange="typechange()">
                            <option value="1" selected>����</option>
                            <option  value="2">����ID��</option>							
                          </select>                          
                ��������				
                <select name="se_ClassName" size="1" id="se_ClassName" onChange="dlchange(document.postart.se_ClassName.options[document.postart.se_ClassName.selectedIndex].value)">
                            <option value="" selected>ȫ�����</option>
                            <%set rs=conn.execute("select * From DNJY_userNewsClass")
							if rs.eof or rs.bof then
							response.write "<option value=''>û�з���</option>"
							else
							do until rs.eof
							%>
                            <option value="<%=rs("classname")%>"><%=rs("ClassName")%></option>
                            <%rs.movenext
							loop							
							end if
							rs.close
							set rs = nothing%>
                </select>

              <div  style="display:inline" id="gjz">�ؼ���&nbsp;��</div><div  style="display:none" id="idh">ID &nbsp;��&nbsp;��</div>
             <input name="searchkeys" type="text" id="searchkeys" size="40">
            <input style="height:20px;" name="Submit" type="submit" class="button" value="�ύ">
            </form>            </td>
          </tr>          
        </table>
</TD>
  </TR>
</table>
  <table width="90%" border="0" cellspacing="1" align="center" class="list">
    <tr>
      <td><font color="#ff0000">ע�⣺Ҫ���е��������ͨ���Ļ�Ա���ܷ������¡����±����ڵ��̿���״̬�²�����ǰ̨��ʾ��</font></TD>
  </TR>
</table>
 <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%">
  <form method="POST" action="UserNews_artplDel.asp?isse=<%=isse%>&classname=<%=classname%>&page=<%=request("page")%>&searchkeys=<%=searchkeys%>&selecttype=<%=selecttype%>&se_ClassName=<%=se_ClassName%>">
    <TR> 
    <TD height="5" colspan="13" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">���е��̵Ļ�Ա���¹���</FONT></TD>
  </TR>
	
    <tr>
      <td width="20" >&nbsp;</td>
      <td width="55" ><div align="center">ID��</div></td>
	  <td width="50" ><div align="center">������Ա</div></td>
      <td width="250"><div align="center">����</div></td>
      <td width="50" ><div align="center">�Ƽ�</div></td>
	  <td width="50" ><div align="center">�̶�</div></td>
      <td width="50" ><div align="center">����</div></td>
	  <td width="55" ><div align="center">�������</div></td>
      <td width="50" ><div align="center">�������</div></td>
      <td width="50" ><div align="center">�����</div></td>
      <td width="88"><div align="center">�������</div></td>
      <td width="88"><div align="center">����</div></td>
      <td width="60"><div align="center">����</div></td>
    </tr>
<%Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1
if rs.eof and rs.bof then
    response.Write("<tr><td colspan='8'><div align='center'>û�м�¼</div> </td></tr>")
else
	rs.PageSize=20
	page=clng(request("page"))
	if page=0 then page=1 
	pages=rs.pagecount
	if page > pages then page=pages
	rs.AbsolutePage=page 
	pp=1
	for j=1 to rs.PageSize	
	Title = rs("Title")
	turl="UserNews_artplDel.asp?id="&rs("id")
	
	modname="UserNews_Modify.asp"
	delname="UserNews_Newsdel.asp"
	if isse=0 then
	caozuocs="isse=0&id="&rs("id")&"&page="&page&"&ClassName="&ClassName
	end if
	if isse=1 then
		 if selecttype=1 then
			caozuocs="isse=1&id="&rs("id")&"&page="&page&"&selecttype="&selecttype&"&searchkeys="&searchkeys&"&se_ClassName="&se_ClassName
		 end if
		 if selecttype=2 then
			caozuocs="isse=1&id="&rs("id")&"&page="&page&"&selecttype="&selecttype&"&searchkeys="&searchkeys
		 end If	 
	end if
%>
	
    
    <tr>
      <td width="20"><input type="checkbox" name="pldel"  value="<%=rs("id")%>"></td>
	  <td width="55"><div align="center"><%=rs("id")%></div></td>
	  <td width="55"><div align="center"><%=rs("username")%></div></td>      
      <td ><a  href="user/article.asp?id=<%=rs("id")%>&com=<%=rs("username")%>"  target=_blank><%=Title%></a></td>
      <td width="58" >
	    <div align="center">
	      <%
		if rs("tj")=true then 
			response.Write "<a href=UserNews_sftj.asp?sf=0&"&caozuocs&"><font color=""#339933"">��</font></a>" 
		else
			response.Write "<a href=UserNews_sftj.asp?sf=1&"&caozuocs&"><font color=""#CC0000"">��</font></a>" 
		end if
		%>
        </div></td>
       <td width="50" >
	    <div align="center">
	      <%
		if rs("newstop")=1 then 
			response.Write "<font color=""#339933"">��</font>" 
		else
			response.Write "<font color=""#CC0000"">��</font>" 
		end if
		%>
        </div></td>
      <td width="50" >
	    <div align="center">
	      <%
		if rs("newsShared")=0 then 
			response.Write "<font color=""#339933"">��</font>" 
		else
			response.Write "<font color=""#CC0000"">��</font>" 
		end if
		%>
        </div></td>
      <td width="0" >
	    <div align="center">
	      <%
		if rs("ifshow")=1 then 
			response.Write "<font color=""#339933"">��</font>" 
		else
			response.Write "<font color=""#CC0000"">��</font>" 
		end if
		%>
        </div></td>
      <td width="0" >
	    <div align="center">
	      <%
		if conn.Execute("Select sdkg From DNJY_user Where username='"&rs("username")&"'")(0)=1 then 
			response.Write "<font color=""#339933"">��</font>" 
		else
			response.Write "<font color=""#CC0000"">��</font>" 
		end if
		%>
        </div></td>
		<td width="58" ><div align="center"><%=rs("hits")%></div></td>
      <td width="88" ><div align="center"><%=rs("classname")%></div></td>
      <td width="88" ><div align="center"><%=rs("addtime")%></div></td>
      <td width="97" ><div align="center"><a href='<%=modname%>?<%=caozuocs%>'>�޸�</a> <a href="javascript:go('<%=delname%>?<%=caozuocs%>','��ȷ��Ҫɾ����')">ɾ��</a>  </div></td>
    </tr>
<%
rs.movenext
pp=pp+1
if rs.eof then exit for
next%>	
	
    
	
    <tr>
      <td colspan="13" > <br> <div align="left">        
          <label>
            <input onclick=CheckAll(this.form) type="checkbox" name="chkall" value="checkbox">
            ȫѡ</label>		  
          <input style="height:20px;" name="Submit" type="submit" class="button" value="����ɾ��">
        
      </div>    </form>    
	  
	  <%
 submitname="UserNews_Modifylist.asp"
 if isse=0 then
 pagecanshu="isse=0&ClassName="&ClassName
 end if
 if isse=1 then
	 if selecttype=1 then
	 	pagecanshu="isse=1&selecttype="&selecttype&"&searchkeys="&searchkeys&"&se_ClassName="&se_ClassName
	 end if
	 if selecttype=2 then
	 	pagecanshu="isse=1&selecttype="&selecttype&"&searchkeys="&searchkeys
	 end If 
 end if
 'pagecanshu="selecttype="&selecttype&"&searchkeys="&searchkeys&"&se_ClassName="&se_ClassName
 %>
 <div align="center">
 <form method=Post action="<%=submitname&"?"&pagecanshu%>">
 <%if Page<2 then      
	response.write "��ҳ ��һҳ&nbsp;"
 else
    response.write "<a href="&submitname&"?"&pagecanshu&"&page=1>��ҳ</a>&nbsp;"
    response.write "<a href="&submitname&"?"&pagecanshu&"&page="&(Page-1)&">��һҳ</a>&nbsp;"
 end if
 if rs.pagecount-page<1 then
    response.write "��һҳ βҳ"
 else
    response.write "<a href="&submitname&"?"&pagecanshu&"&page="&(page+1)&">"
    response.write "��һҳ</a> <a href="&submitname&"?"&pagecanshu&"&page="&rs.pagecount&">βҳ</a>"
  end if
   response.write "&nbsp;ҳ�Σ�<strong><font color=red>"&Page&"</font>/"&rs.pagecount&"</strong>ҳ "
   response.write "&nbsp;��<b><font color='#FF0000'>"&rs.recordcount&"</font></b>����¼ <b>"&rs.pagesize&"</b>����¼/ҳ"
   response.write " ת����<input type='text' name='page' size=4 maxlength=10  value="&page&">"
   response.write " <input  type='submit'  class='button' value=' Goto '  name='cndok'></span>"     
%>
 </form>
 </div>
      </td>      
    </tr>
<%
end if
set rs=Nothing
Call CloseDB(conn)
%>
	</table>
	</td>
	  </tr>
  </table>
</body>

</html>
