<!--#include file="sdtop.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
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
        <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td><div id="lt2"><span style="margin-left:50px;">��������</span></div></td>
          </tr>
        </table>
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="35%">
<table align='center' cellpadding='0' cellspace='0' border='0'>
<tr>
	<td id='demo1' valign='top'>
<div align="center">
	<TABLE border=0 width="100%">
<tr>
<%dim j,rstj_a,sqltj_a,catid

				Const MaxPerPage=20 'ÿҳ��ʾ����
   				dim totalPut   
   				dim CurrentPage
   				dim TotalPages   				
    			if Not isempty(Request("page")) then
      				currentPage=Cint(Request("page"))
   				else
      				currentPage=1
   				end if 
set rstj_a=server.createobject("adodb.recordset")
sqltj_a="select * from DNJY_comlink where username='"&com&"' order by ID "&DN_OrderType&"" 	      
rstj_a.open sqltj_a,conn,1,1
				
  				if rstj_a.eof And rstj_a.bof Then				
       			Response.Write "<td align='center'><p align='center' class='contents'> ��Ǹ����ʱ��û�����ӣ�</p>"
				response.write "</td></tr></table></td></tr></table>"
   				else
	  				totalPut=rstj_a.recordcount

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
            			showpage totalput,MaxPerPage,"userlink_list.asp?com="&com&"&"&C_ID&""
       				else
          				if (currentPage-1)*MaxPerPage<totalPut then
            				rstj_a.move  (currentPage-1)*MaxPerPage
            				dim bookmark
            				bookmark=rstj_a.bookmark
            				showContent
             				showpage totalput,MaxPerPage,"userlink_list.asp?com="&com&"&"&C_ID&""
        				else
	        				currentPage=1
           					showContent
           					showpage totalput,MaxPerPage,"userlink_list.asp?com="&com&"&"&C_ID&""
	      				end if
	   				end if
   				end if

   				sub showContent
				dim i
	   			i=0               
do while not rstj_a.eof                         
  %> 
<td align="center"><a target="_blank" href="<%=rstj_a("url")%>"><u><b><%=rstj_a("web")%></b></u></a>&nbsp;&nbsp;</td>
<%     
i=i+1    
	if i>=MaxPerPage then exit Do	'ÿҳ��ʾ������������MaxPerPage�趨����
    rstj_a.movenext
    loop
    rstj_a.close
   set rstj_a=nothing   
%> 
</TABLE></div>
	</td>
	</tr>
</table>
   <%  
				End Sub   
  
				Function showpage(totalnumber,maxperpage,filename)  
  				Dim n
  				
				If totalnumber Mod maxperpage=0 Then  
					n= totalnumber \ maxperpage  
				Else
					n= totalnumber \ maxperpage+1  
				End If %>
                  <form method=Post action= <% = filename %>&catid=<% = catid %>>
                    <p align='center'> 
                      <%
				If CurrentPage<2 Then  %>
                      ��ҳ ��ҳ 
                      <%
				Else   %>
                      <a href=<% = filename %>&page=1&catid=<% = catid %>&<%=C_ID%>>��ҳ</a> 
                      <a href=<% = filename %>&page=<% = CurrentPage-1 %>&catid=<% = catid %>&<%=C_ID%>>��ҳ</a> 
                      <%
				End If	
				If n-currentpage<1 Then  %>
                      ��ҳ ĩҳ 
                      <%
				Else   %>
                      <a href=<% = filename %>&page=<% = (CurrentPage+1) %>&catid=<% = catid %>&<%=C_ID%>>��ҳ</a> 
                      <a href=<% = filename %>&page=<% = n %>&catid=<% = catid %>&<%=C_ID%>>ĩҳ</a> 
                      <%
				End If  %>
                      �� <b> 
                      <% = CurrentPage %>
                      </b> ҳ �� <b> 
                      <% = n %>
                      </b> ҳ ��&nbsp;<b> 
                      <% = totalnumber %>
                      </b>&nbsp;�����ӡ�ÿҳ <b> 
                      <% = maxperpage %>
                      </b> ������ ת���ڣ� 
                      <input type='text' name='page' size=2 maxlength=10 class=smallInput value=<% = currentpage %>>
                      ҳ 
                      <input type='submit'  class='contents' value='��ת' name='cndok'>
                  </form>
                  <%
				End Function  
			%>
<!--��Ʒ�������--></td>
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
              <td align="center"><script language=javascript src="<%=adjs_path("ads/js","xxfl1",c1&"_"&c2&"_"&c3)%>"></script></td>
            </tr>
          </table><%End if%>		

</td>
</tr>			  
</table>
<!--#include file="sdend.asp"-->