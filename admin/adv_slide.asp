<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<!--#include file=../ads/Html2Js.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("13")%>
<html>
<head>
<title>�õƹ��A</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<link href="../css/style.css" rel="stylesheet" type="text/css">
 <script language="javascript">
<!--
function form_onsubmit() {
if (document.form.pic1.value=="")
	{
	  alert("���ϴ�ͼƬһ��")
	  document.form.pic1.focus()
	  return false
	 }

if (document.form.pic1_lnk.value=="")
	{
	  alert("����дͼƬһ���ӣ�")
	  document.form.pic1_lnk.focus()
	  return false
	 }
if (document.form.tit1.value=="" )
	{
	  alert("����ͼƬһ˵����")
	  document.form.tit1.focus()
	  return false
	 }
if (document.form.pic2.value=="")
	{
	  alert("���ϴ�ͼƬ����")
	  document.form.pic2.focus()
	  return false
	 }

if (document.form.pic2_lnk.value=="")
	{
	  alert("����дͼƬ�����ӣ�")
	  document.form.pic2_lnk.focus()
	  return false
	 }
if (document.form.tit2.value=="" )
	{
	  alert("����ͼƬ��˵����")
	  document.form.tit2.focus()
	  return false
	 }
if (document.form.pic3.value=="")
	{
	  alert("���ϴ�ͼƬ����")
	  document.form.pic3.focus()
	  return false
	 }

if (document.form.pic3_lnk.value=="")
	{
	  alert("����дͼƬ�����ӣ�")
	  document.form.pic3_lnk.focus()
	  return false
	 }
if (document.form.tit3.value=="" )
	{
	  alert("����ͼƬ��˵����")
	  document.form.tit3.focus()
	  return false
	 } 	 
if (document.form.width.value=="" )
	{
	  alert("����дͼƬ��ȣ�")
	  document.form.width.focus()
	  return false
	 }		 
if (document.form.height.value=="" )
	{
	  alert("����дͼƬ�߶ȣ�")
	  document.form.height.focus()
	  return false
	 }		 
 
}
// --></script>
<script language="JavaScript" type="text/JavaScript">
function uppic1(model,frmname) {
//ȷ���ϴ��ļ�������ѡ�������������жϱ�Ų��Ա������
id=document.form.ADID.value;
id1=document.form.city_oneid.value;
id2=document.form.city_twoid.value;
id3=document.form.city_threeid.value;
if (id1=='') id1=0
if (id2=='') id2=0
if (id3=='') id3=0
if (id=='') {
 alert("������д���ID�ţ�");
 document.form.ADID.focus();
return false;}
window.open("../ads/DNJY_upload_ad.asp?ttup=slide1&fupname="+id+model+"_t1"+"_"+id1+"_"+id2+"_"+id3+"&frmname="+frmname,"blank_","scrollbars=yes,resizable=no,top=40,left=150,width=500,height=150")
}
function uppic2(model,frmname) {
//ȷ���ϴ��ļ�������ѡ�������������жϱ�Ų��Ա������
id=document.form.ADID.value;
id1=document.form.city_oneid.value;
id2=document.form.city_twoid.value;
id3=document.form.city_threeid.value;
if (id1=='') id1=0
if (id2=='') id2=0
if (id3=='') id3=0
if (id=='') {
 alert("������д���ID�ţ�");
 document.form.ADID.focus();
return false;}
window.open("../ads/DNJY_upload_ad.asp?ttup=slide2&fupname="+id+model+"_t2"+"_"+id1+"_"+id2+"_"+id3+"&frmname="+frmname,"blank_","scrollbars=yes,resizable=no,top=40,left=150,width=500,height=150")
}
function uppic3(model,frmname) {
//ȷ���ϴ��ļ�������ѡ�������������жϱ�Ų��Ա������
id=document.form.ADID.value;
id1=document.form.city_oneid.value;
id2=document.form.city_twoid.value;
id3=document.form.city_threeid.value;
if (id1=='') id1=0
if (id2=='') id2=0
if (id3=='') id3=0
if (id=='') {
 alert("������д���ID�ţ�");
 document.form.ADID.focus();
return false;}
window.open("../ads/DNJY_upload_ad.asp?ttup=slide3&fupname="+id+model+"_t3"+"_"+id1+"_"+id2+"_"+id3+"&frmname="+frmname,"blank_","scrollbars=yes,resizable=no,top=40,left=150,width=500,height=150")
}
</script>
</head>
<BODY>
	<table width="98%" border="1"  style="border-collapse: collapse; border-style: dotted; border-width: 0px"  bordercolor="#333333" cellspacing="0" cellpadding="2">

<%dim action,pic,n,rscity,id,id1,id2,id3
action=request("ok")
id=request.Form("ADID")
id1=request.Form("city_oneid")
id2=request.Form("city_twoid")
id3=request.Form("city_threeid")
if id1="" Then id1=0
if id2="" Then id2=0
if id3="" Then id3=0
Call OpenConn%>	
	<tr class=backs><td colspan=3 class=td height=18 align="right">
<form action=adv_slide.asp method=post name=myform>	
	�����������õƹ��A���б༭��
<%
set rscity=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")%>
          <SCRIPT language = "JavaScript">
var onecount1;
onecount1=0;
subcat1 = new Array();
        <%
	dim count1:count1 = 0
        do while not rscity.eof 
        %>
subcat1[<%=count1%>] = new Array("<%=rscity("city")%>","<%=rscity("id")%>","<%=rscity("twoid")%>");
        <%
        count1 = count1 + 1
        rscity.movenext
        loop
        rscity.close
		set rscity=nothing
        %>
onecount1=<%=count1%>;
          </SCRIPT>
          <%
set rscity=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid>0")
%>
          <SCRIPT language = "JavaScript">
var onecount2;
onecount2=0;
subcat2 = new Array();
        <%
		dim count2:count2 = 0
        do while not rscity.eof 
        %>
subcat2[<%=count2%>] = new Array("<%=rscity("city")%>","<%=rscity("id")%>","<%=rscity("twoid")%>","<%=rscity("threeid")%>");
        <%
        count2 = count2 + 1
        rscity.movenext
        loop
        rscity.close
		set rscity = nothing
        %>
onecount2=<%=count2%>;

function changelocation1(locationid1)
    {
    document.myform.city_twoid.length = 0; 
	document.myform.city_twoid.options[0] = new Option('ѡ�����','');
	document.myform.city_threeid.length = 0; 
	document.myform.city_threeid.options[0] = new Option('ѡ�����','');
    var locationid1=locationid1;
    var i;
    for (i=0;i < onecount1; i++)
        {
            if (subcat1[i][1] == locationid1)
            { 
                document.myform.city_twoid.options[document.myform.city_twoid.length] = new Option(subcat1[i][0], subcat1[i][2]);
            }        
        }
        
    }    
	
	function changelocation2(locationid1,locationid2)
    {
    document.myform.city_threeid.length = 0; 
    document.myform.city_threeid.options[0] = new Option('ѡ�����','');
    var locationid1=locationid1;
	 var locationid2=locationid2;
    var i;
    for (i=0;i < onecount2; i++)
        {
            if (subcat2[i][2] == locationid1)
            { 
			if (subcat2[i][1] == locationid2)
			{
                document.myform.city_threeid.options[document.myform.city_threeid.length] = new Option(subcat2[i][0], subcat2[i][3]);
            }        
        }
       } 
    }            
              </SCRIPT>
          <SELECT name="city_oneid" size="1" id="select" onChange="changelocation1(document.myform.city_oneid.options[document.myform.city_oneid.selectedIndex].value)">
            <OPTION value="0" selected>ѡ�����</OPTION>
            <%set rscity=conn.execute("select * from DNJY_city where id>0 and twoid=0 order by indexid asc")
if rscity.eof or rscity.bof then
response.write "<option value='0'>û�з���</option>"
else
do until rscity.eof
response.write "<option value='"&rscity("id")&"'>"&rscity("city")&"</option>"
    rscity.movenext
    loop
end if
rscity.close:set rscity = nothing
%>
          </SELECT>
          <SELECT name="city_twoid" id="select1" onChange="changelocation2(document.myform.city_twoid.options[document.myform.city_twoid.selectedIndex].value,document.myform.city_oneid.options[document.myform.city_oneid.selectedIndex].value)">
            <OPTION value="0" selected>ѡ�����</OPTION>
          </SELECT>
          <SELECT name="city_threeid" id="select2">
            <OPTION value="0" selected>ѡ�����</OPTION>
          </SELECT> 
<INPUT name="ok" TYPE="hidden" value="search"><INPUT name=action TYPE="submit" value="����">
</form>&nbsp;&nbsp;&nbsp;&nbsp;�����Ҫ�����༭��վ�õƣ���ѡ�������ɣ�</td>	
	
	</tr>
	<tr class=backs><td colspan=3 class=td height=18 align="center"><b>��ӻõƹ��A����ҳ�õƹ��AΪ�ֵ����������ʾ���ڻ��������п�������ʾA��B��</b></td></tr>
<form action=adv_slide.asp method=post name=form onSubmit="return form_onsubmit()">
<%if action="ok" Or action="edit" then
	if trim(request.form("pic1"))="" or trim(request.form("pic2"))="" or trim(request.form("pic3"))="" then
	response.write "<script language='javascript'>"
	response.write "alert('���Ź���ͼƬ��·�����ļ���������д��������ȷ����������һҳ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	pic=request.form("pic1")
	pic=split(pic,".")
	n=UBound(pic)
	if LCase(pic(n))<>"jpg" then
	response.write "<script language='javascript'>"
	response.write "alert('ͼƬһ�ĸ�ʽ����ȷ����֧��JPG��������ȷ����������һҳ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	pic=request.form("pic2")
	pic=split(pic,".")
	n=UBound(pic)
	if LCase(pic(n))<>"jpg" then
	response.write "<script language='javascript'>"
	response.write "alert('ͼƬһ�ĸ�ʽ����ȷ����֧��JPG��������ȷ����������һҳ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	pic=request.form("pic3")
	pic=split(pic,".")
	n=UBound(pic)
	if LCase(pic(n))<>"jpg" then
	response.write "<script language='javascript'>"
	response.write "alert('ͼƬһ�ĸ�ʽ����ȷ����֧��JPG��������ȷ����������һҳ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if
end if

if action="" Then%>
	<tr><td width=20% align=right height="18">����������</td><td>
<%
set rscity=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")%>
          <SCRIPT language = "JavaScript">
var onecount5;
onecount5=0;
subcat5 = new Array();
        <%
	dim count5:count5 = 0
        do while not rscity.eof 
        %>
subcat5[<%=count5%>] = new Array("<%=rscity("city")%>","<%=rscity("id")%>","<%=rscity("twoid")%>");
        <%
        count5 = count5 + 1
        rscity.movenext
        loop
        rscity.close
		set rscity=nothing
        %>
onecount5=<%=count5%>;
          </SCRIPT>
          <%
set rscity=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid>0")
%>
          <SCRIPT language = "JavaScript">
var onecount4;
onecount4=0;
subcat4 = new Array();
        <%
		dim count4:count4 = 0
        do while not rscity.eof 
        %>
subcat4[<%=count4%>] = new Array("<%=rscity("city")%>","<%=rscity("id")%>","<%=rscity("twoid")%>","<%=rscity("threeid")%>");
        <%
        count4 = count4 + 1
        rscity.movenext
        loop
        rscity.close
		set rscity = nothing
        %>
onecount4=<%=count4%>;

function changelocation(locationid)
    {
    document.form.city_twoid.length = 0; 
	document.form.city_twoid.options[0] = new Option('ѡ�����','');
	document.form.city_threeid.length = 0; 
	document.form.city_threeid.options[0] = new Option('ѡ�����','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount5; i++)
        {
            if (subcat5[i][1] == locationid)
            { 
                document.form.city_twoid.options[document.form.city_twoid.length] = new Option(subcat5[i][0], subcat5[i][2]);
            }        
        }
        
    }    
	
	function changelocation4(locationid,locationid1)
    {
    document.form.city_threeid.length = 0; 
    document.form.city_threeid.options[0] = new Option('ѡ�����','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount4; i++)
        {
            if (subcat4[i][2] == locationid)
            { 
			if (subcat4[i][1] == locationid1)
			{
                document.form.city_threeid.options[document.form.city_threeid.length] = new Option(subcat4[i][0], subcat4[i][3]);
            }        
        }
       } 
    }            
              </SCRIPT>
          <SELECT name="city_oneid" size="1" id="select5" onChange="changelocation(document.form.city_oneid.options[document.form.city_oneid.selectedIndex].value)">
            <OPTION value="0" selected>ѡ�����</OPTION>
            <%set rscity=conn.execute("select * from DNJY_city where id>0 and twoid=0 order by indexid asc")
if rscity.eof or rscity.bof then
response.write "<option value='0'>û�з���</option>"
else
do until rscity.eof
response.write "<option value='"&rscity("id")&"'>"&rscity("city")&"</option>"
    rscity.movenext
    loop
end if
rscity.close:set rscity = nothing
%>
          </SELECT>
          <SELECT name="city_twoid" id="select6" onChange="changelocation4(document.form.city_twoid.options[document.form.city_twoid.selectedIndex].value,document.form.city_oneid.options[document.form.city_oneid.selectedIndex].value)">
            <OPTION value="0" selected>ѡ�����</OPTION>
          </SELECT>
          <SELECT name="city_threeid" id="select7">
            <OPTION value="0" selected>ѡ�����</OPTION>
          </SELECT>&nbsp;&nbsp;&nbsp;�������ѡ���������Ϊ��վ��档���ԭ���Ѵ��ڣ������������༭���ɣ�</td></tr>
	<tr><td width=20% align=right height="18">ͼƬһ��·�����ļ�����</td><td> <input type=text value="" name=pic1 size=40 maxlength=100>
	<INPUT  TYPE="button" value="���ͼƬ�ϴ�" onclick="javascript:uppic1('','pic1');"> ����jpg��ʽͼƬ������510*280���أ�500K��</td></tr>	
	<tr><td width=20% align=right height="18">���ӣ�</td><td> <input type=text value="" name=pic1_lnk size=40 maxlength=100></td></tr>		
	<tr><td width=20% align=right height="18">˵����</td><td> <input type=text value="" name=tit1 size=40 maxlength=100></td></tr>
	
	<tr><td width=20% align=right height="18">ͼƬ����·�����ļ�����</td><td> <input type=text value="" name=pic2 size=40 maxlength=100>
	<INPUT  TYPE="button" value="���ͼƬ�ϴ�" onclick="javascript:uppic2('','pic2');"> ����jpg��ʽͼƬ������510*280���أ�500K��</td></tr>	
	<tr><td width=20% align=right height="18">���ӣ�</td><td> <input type=text value="" name=pic2_lnk size=40 maxlength=100></td></tr>			
	<tr><td width=20% align=right height="18">˵����</td><td> <input type=text value="" name=tit2 size=40 maxlength=100></td></tr>
	
	<tr><td width=20% align=right height="18">ͼƬ����·�����ļ�����</td><td><input type=text value="" name=pic3 size=40 maxlength=100>
	<INPUT  TYPE="button" value="���ͼƬ�ϴ�" onclick="javascript:uppic3('','pic3');"> ����jpg��ʽͼƬ������510*280���أ�500K��</td></tr>
	<tr><td width=20% align=right height="18">���ӣ�</td><td> <input type=text value="" name=pic3_lnk size=40 maxlength=100></td></tr>
	<tr><td width=20% align=right height="18">˵����</td><td> <input type=text value="" name=tit3 size=40 maxlength=100></td></tr>		    
	
	<tr><td width=20% align=right height="18"></td></tr>	
	<tr><td width=20% align=right height="18">ͼƬ��ȣ�</td><td colspan=3> <input type=text value="510"  name=width size=40 maxlength=100> ����ͼƬ�Ŀ�ȣ�ҪΪ��������������510</td></tr>
	<tr><td width=20% align=right height="18">ͼƬ�߶ȣ�</td><td colspan=3> <input type=text value="280"  name=height size=40 maxlength=100>  ����ͼƬ�ĸ߶ȣ�ҪΪ��������������280</td></tr>
	<tr><td width=20% align=right height="18">���ÿ��أ�</td><td>            
                <input type="radio" name="switcha" value="1" id="switcha" checked>
                 ����&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="switcha" value="0" id="switcha">
                �ر�</td></tr>
	<input type="hidden" name="ADID" value="slide">
	<tr><td colspan=3>
<INPUT name="ok" TYPE="hidden" value="ok"><INPUT name=action TYPE="submit" value="��Ӳ�����JS�ļ�"></td></tr>
</form></table>
<%End if

if action="search" Then
Dim TempText
Set rs = conn.Execute("select * from DNJY_slide where city_oneid="&id1&" and city_twoid="&id2&" and city_threeid="&id3&"") 
if rs.eof and rs.bof Then
TempText=TempText&"<tr><td><p><br>����"
IF id1>0 Then TempText=TempText&""& conn.Execute("Select city From DNJY_city Where id="&id1&" and twoid="&id2&" and threeid="&id3&"")(0)&""
response.write TempText&"�õƹ�棬��<a href=""adv_slide.asp"" TARGET=right><font color=""#0000ff""> ��ӻõƹ��A</font></a></td></tr>"
response.end
else'����м�¼%>
	<tr><td width=20% align=right height="18">����������</td><td>
<%
set rscity=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")%>
<SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
		dim count:
		count = 0
        do while not rscity.eof 
        %>
subcat[<%=count%>] = new Array("<%=rscity("city")%>","<%=rscity("id")%>","<%=rscity("twoid")%>");
        <%
        count = count + 1
        rscity.movenext
        loop
        rscity.close
		set rscity=nothing
        %>
onecount=<%=count%>;
                             </SCRIPT>
                                   <%
set rscity=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid>0")
%>
                                   <SCRIPT language = "JavaScript">
var onecount4;
onecount4=0;
subcat4 = new Array();
        <%
		count4 = 0
        do while not rscity.eof 
        %>
subcat4[<%=count4%>] = new Array("<%=rscity("city")%>","<%=rscity("id")%>","<%=rscity("twoid")%>","<%=rscity("threeid")%>");
        <%
        count4 = count4 + 1
        rscity.movenext
        loop
        rscity.close
		set rscity = nothing
        %>
onecount4=<%=count4%>;

function changelocation(locationid)
    {
    document.form.city_twoid.length = 0; 
	document.form.city_twoid.options[0] = new Option('ѡ�����','');
	document.form.city_threeid.length = 0; 
	document.form.city_threeid.options[0] = new Option('ѡ�����','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.form.city_twoid.options[document.form.city_twoid.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
	
	function changelocation4(locationid,locationid1)
    {
    document.form.city_threeid.length = 0; 
    document.form.city_threeid.options[0] = new Option('ѡ�����','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount4; i++)
        {
            if (subcat4[i][2] == locationid)
            { 
			if (subcat4[i][1] == locationid1)
			{
                document.form.city_threeid.options[document.form.city_threeid.length] = new Option(subcat4[i][0], subcat4[i][3]);
            }        
        }
       } 
    }            
                             </SCRIPT>
                                   <SELECT name="city_oneid" size="1" id="select2" onChange="changelocation(document.form.city_oneid.options[document.form.city_oneid.selectedIndex].value)">
                                     <OPTION value="" selected>ѡ�����</OPTION>
                                     <%set rscity=conn.execute("select * from DNJY_city where id>0 and twoid=0 order by indexid asc")
if rscity.eof and rscity.bof then
response.write "<option value=''>û�з���</option>"
else
do until rscity.eof%>
                                     <OPTION value="<%=rscity("id")%>" <%if rscity("id")=rs("city_oneid") then%>selected<%end if%>><%=rscity("city")%></OPTION>
                                     <% rscity.movenext
    loop
end if
rscity.close
set rscity = nothing
%>
                                   </SELECT>
                                   <SELECT name="city_twoid" id="city_twoid" onChange="changelocation4(document.form.city_twoid.options[document.form.city_twoid.selectedIndex].value,document.form.city_oneid.options[document.form.city_oneid.selectedIndex].value)">
                                     <OPTION value="" selected>ѡ�����</OPTION>
                                     <%set rscity=conn.execute("select * from DNJY_city where id="&rs("city_oneid")&" and twoid<>0 and threeid=0")
if rscity.eof and rscity.bof then
response.write "<option value=''>û�з���</option>"
else
do until rscity.eof%>
                                     <OPTION value="<%=rscity("twoid")%>" <%if rscity("twoid")=rs("city_twoid") then%>selected<%end if%>><%=rscity("city")%></OPTION>
                                     <% rscity.movenext
    loop
	end if
rscity.close
set rscity = nothing%>
                                   </SELECT>
                                   <SELECT name="city_threeid" id="city_threeid">
                                     <OPTION value="" selected>ѡ�����</OPTION>
                                     <%set rscity=conn.execute("select * from DNJY_city where id="&rs("city_oneid")&" and twoid="&rs("city_twoid")&" and threeid<>0")
if rscity.eof and rscity.bof then
response.write "<option value=''>û�з���</option>"
else
do until rscity.eof%>
                                     <OPTION value="<%=rscity("threeid")%>" <%if rscity("threeid")=rs("city_threeid") then%>selected<%end if%>><%=rscity("city")%></OPTION>
                                     <% rscity.movenext
    loop
	end if
rscity.close
set rscity = nothing%>
                                   </SELECT> �����ѡ���������ʾ��վ���</td> 
	<td  align="center"></td></tr>
	<tr><td width=20% align=right height="18">ͼƬһ��·�����ļ�����</td><td> <input type=text value="<%=rs("pic1")%>" name=pic1 size=40 maxlength=100>
	<INPUT  TYPE="button" value="���ͼƬ�ϴ�" onclick="javascript:uppic1('','pic1');"> ����jpg��ʽͼƬ��
    <%=rs("width")%>*<%=rs("height")%>���أ�500K��</td><td  align="center" rowspan="3"><a href=<%=rs("pic1")%> target=_blank><img src=<%=rs("pic1")%> border=0  alt="���Ԥ��Ч��" width=180 height=110></a>
	</td></tr>	
	<tr><td width=20% align=right height="18">���ӣ�</td><td> <input type=text value="<%=rs("pic1_lnk")%>" name=pic1_lnk size=40 maxlength=100></td></tr>		
	<tr><td width=20% align=right height="18">˵����</td><td> <input type=text value="<%=rs("tit1")%>" name=tit1 size=40 maxlength=100></td></tr>	
	<tr><td width=20% align=right height="18">ͼƬ����·�����ļ�����</td><td> <input type=text value="<%=rs("pic2")%>" name=pic2 size=40 maxlength=100>
	<INPUT  TYPE="button" value="���ͼƬ�ϴ�" onclick="javascript:uppic2('','pic2');"> ����jpg��ʽͼƬ��<%=rs("width")%>*<%=rs("height")%>���أ�500K��
	</td>
	<td  align="center" rowspan="3"><a href=<%=rs("pic2")%> target=_blank><img src=<%=rs("pic2")%> border=0  alt="���Ԥ��Ч��" width=180 height=110></a>
	</td>	</tr>	
	<tr><td width=20% align=right height="18">���ӣ�</td><td> <input type=text value="<%=rs("pic2_lnk")%>" name=pic2_lnk size=40 maxlength=100></td></tr>			
	<tr><td width=20% align=right height="18">˵����</td><td> <input type=text value="<%=rs("tit2")%>" name=tit2 size=40 maxlength=100></td></tr>	
	<tr><td width=20% align=right height="18">ͼƬ����·�����ļ�����</td><td><input type=text value="<%=rs("pic3")%>" name=pic3 size=40 maxlength=100>
	<INPUT  TYPE="button" value="���ͼƬ�ϴ�" onclick="javascript:uppic3('','pic3');"> ����jpg��ʽͼƬ��<%=rs("width")%>*<%=rs("height")%>���أ�500K��
	</td>
	<td  align="center" rowspan="3"><a href=<%=rs("pic3")%> target=_blank><img src=<%=rs("pic3")%> border=0  alt="���Ԥ��Ч��" width=180 height=110></a>
	</td>	</tr>
	<tr><td width=20% align=right height="18">���ӣ�</td><td> <input type=text value="<%=rs("pic3_lnk")%>" name=pic3_lnk size=40 maxlength=100></td></tr>
	<tr><td width=20% align=right height="18">˵����</td><td> <input type=text value="<%=rs("tit3")%>" name=tit3 size=40 maxlength=100></td></tr>		    
	
	<tr><td width=20% align=right height="18"></td></tr>	
	<tr><td width=20% align=right height="18">ͼƬ��ȣ�</td><td colspan=3> <input type=text value="<%=rs("width")%>"  name=width size=40 maxlength=100> ����ͼƬ�Ŀ�ȣ�ҪΪ��������������510</td></tr>
	<tr><td width=20% align=right height="18">ͼƬ�߶ȣ�</td><td colspan=3> <input type=text value="<%=rs("height")%>"  name=height size=40 maxlength=100>  ����ͼƬ�ĸ߶ȣ�ҪΪ��������������208</td></tr>
	<tr><td width=20% align=right height="18">���ÿ��أ�</td><td> <%if rs("switcha")=1 then%>               
                <input type="radio" name="switcha" value="1" id="switcha" checked>
                 ����&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="switcha" value="0" id="switcha">
                �ر�
                <%else%>                         
                <input type="radio" name="switcha" value="1" id="switcha">
                 ����&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="switcha" value="0" id="switcha" checked>
                �ر�<%end if%></td></tr>
	<input type="hidden" name="ADID" value="slide">
	<tr><td colspan=3>
<INPUT name="ok" TYPE="hidden" value="edit"><INPUT name=action TYPE="submit" value="�������ò�����JS�ļ�"></td></tr>
</form></table>
	<%end If
	 end If
	if action="ok" Or action="edit" then	
     Set rs=Server.CreateObject("ADODB.Recordset")
	 sql="select * from DNJY_slide where city_oneid="&id1&" and city_twoid="&id2&" and city_threeid="&id3&""
	 rs.open sql,conn,1,3
	 If action="ok" Then
	 if Not (rs.eof or rs.bof) then
	 Response.Write ("<script language=javascript>alert('�˵������лõƹ�棬������ӣ�ֻ�������޸�!');history.go(-1);</script>")
     Response.End
     else
	 rs.addnew	
	 End If
	 End If
	 
 	 rs("city_oneid")=id1
 	 rs("city_twoid")=id2 
 	 rs("city_threeid")=id3
 	 rs("adsjs")=""&DangerEncode(Request.Form("ADID"))&"_"&id1&"_"&id2&"_"&id3&".js" 

 	 rs("pic1")=trim(request.form("pic1"))
 	 rs("pic1_lnk")=trim(request.form("pic1_lnk"))
 	 rs("tit1")=trim(request.form("tit1"))

 	 rs("pic2")=trim(request.form("pic2"))
 	 rs("pic2_lnk")=trim(request.form("pic2_lnk"))
 	 rs("tit2")=trim(request.form("tit2"))

 	 rs("pic3")=trim(request.form("pic3"))
 	 rs("pic3_lnk")=trim(request.form("pic3_lnk"))
 	 rs("tit3")=trim(request.form("tit3"))

 	 rs("width")=trim(request.form("width"))
 	 rs("height")=trim(request.form("height"))
	 rs("switcha")=trim(request.form("switcha"))
	rs.update

'����js���뿪ʼ
Dim JsCode
		JsCode="<script type=""text/javascript"">" & vbCrLf & _
		"var focus_width="&rs("width")&";" & vbCrLf & _
		"var focus_height="&rs("height")&";" & vbCrLf & _
		"var text_height=20;" & vbCrLf & _
		"var swf_height = focus_height+text_height;" & vbCrLf & _
		"var pics='',links='', texts='';" & vbCrLf & _
		"pics+='|"&rs("pic1")&"';links+='|"&rs("pic1_lnk")&"';texts+='|"&rs("tit1")&"';" & vbCrLf & _
		"pics+='|"&rs("pic2")&"';links+='|"&rs("pic2_lnk")&"';texts+='|"&rs("tit2")&"';" & vbCrLf & _
		"pics+='|"&rs("pic3")&"';links+='|"&rs("pic3_lnk")&"';texts+='|"&rs("tit3")&"';" & vbCrLf & _
		
		"pics=pics.substring(1);links=links.substring(1);texts=texts.substring(1);" & vbCrLf & _

		"document.write('<object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"" codebase=""http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0"" width=""'+ focus_width +'"" height=""'+ swf_height +'"">');" & vbCrLf & _
		"document.write('<param name=""allowScriptAccess"" value=""sameDomain""><param name=""movie"" value=""/"&strInstallDir&"images/focus_slide.swf""><param name=""quality"" value=""high""><param name=""bgcolor"" value=""#F2F2F2"">');" & vbCrLf & _
		"document.write('<param name=""menu"" value=""false""><param name=wmode value=""opaque"">');" & vbCrLf & _
		"document.write('<param name=""FlashVars"" value=""pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'"">');" & vbCrLf & _
		"document.write('<embed src=""/"&strInstallDir&"images/focus_slide.swf"" wmode=""opaque"" FlashVars=""pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'"" menu=""false"" ?bgcolor=""#ffffff"" quality=""high"" width=""'+ focus_width +'"" height=""'+ swf_height +'"" allowScriptAccess=""sameDomain"" type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" />');  document.write('</object>');" & vbCrLf & _
		"</script>" & vbCrLf & _
""
If trim(request.form("switcha"))=0 Then JsCode="<script type=""text/javascript""> <!--�����ͣ ID:"&ID&"_"&id1&"_"&id2&"_"&id3&"--></script>"
JsCode=Html2Js(JsCode)
JsCode="document.write('"&JsCode&"')"
JsCode = "<!--" & vbCrLf & JsCode & vbCrLf &"//-->" 
dim fs,f
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.CreateTextFile(Server.MapPath("/")&"/"&""&strInstallDir&"UploadFiles/slidea/"&id&"_"&id1&"_"&id2&"_"&id3&".js", True)
f.write(JsCode)
f.close
Set f = nothing
Set fs = Nothing
'����js�������
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
	response.write "<script language='javascript'>"
	response.write "alert('�õ�ͼƬ���óɹ���������JS�ļ���');"
	response.write "location.href='adv_slide.asp';"
	response.write "</script>"
	response.end
end If

 '�ж��Ƿ��վ��ʾ�����Ƿ��й��JS�ļ����ڣ�������ʾ��վ��棽����������������������������������
Function adjs_path(p,adid,jsid)
	Dim TempText,Fso
	TempText=Server.MapPath(p)
	Set Fso = CreateObject("Scripting.FileSystemObject")
	IF Fso.FileExists(TempText&"/"&adid&"_"&jsid&".js") and shows8=1 Then	
	adjs_path=p&"/"&adid&"_"&jsid&".js"		
	Else
	adjs_path=p&"/"&adid&"_0_0_0.js"	
	End IF
	Set Fso =Nothing
End Function
Rem ���˿��ܳ�����ķ���
Function DangerEncode(fString)
If not isnull(fString) Then
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10), "")
	fString = replace(fString, "'", """")
    fString = Trim(fString)
    DangerEncode = fString
End If
End Function%>

</body>
<TABLE align="center">
<TR>
	<TD><button onClick="document.location.reload()">ˢ�¿�Ч��</button>Ч��Ԥ����<br>
<table width="100%" cellspacing="0" cellpadding="0" class="tablest">
<tr><td width="500" ><script src="<%=adjs_path("../UploadFiles/slidea","slide",id1&"_"&id2&"_"&id3)%>"></script></td>
</tr></table>
<%
	Dim c1,c2,c3%></TD>
	<TD>��̬ҳ���÷�ʽ��&lt;script src="&lt;%=adjs_path("UploadFiles/slidea","slide",c1&"_"&c2&"_"&c3)%>">&lt;/script> <br> ��̬ҳ���÷�ʽ��&lt;script src="UploadFiles/slidea/slide_һ������ID��_��������ID��_��������ID��.js">&lt;/script></TD>
</TR>
</TABLE>
<p></html>

