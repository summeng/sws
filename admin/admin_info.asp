<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Call OpenConn
dim ThisPage
id=request("id")
if request("page")="" then
  ThisPage=1		
else
  ThisPage=request("page")
end if

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
Call CloseRs(rs)
	If GetIPAddressData = "" Then GetIPAddressData="��&#1450;"
end function
%>
<script language="JavaScript">
function DrawImage(ImgD){
   var image=new Image();
   image.src=ImgD.src;
   if(image.width>0 && image.height>0){
    flag=true;
    if(image.width/image.height>= 120/80){
     if(image.width>120){  
     ImgD.width=80;
     ImgD.height=(image.height*80)/image.width;
     }else{
     ImgD.width=image.width;  
     ImgD.height=image.height;
     }

     }
    else{
     if(image.height>80){  
     ImgD.height=80;
     ImgD.width=(image.width*80)/image.height;     
     }else{
     ImgD.width=image.width;  
     ImgD.height=image.height;
     }

     }
    }
   /*else{
    ImgD.src="";
    ImgD.alt=""
    }*/
   }
function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.name != 'chkall' )
       e.checked = form.chkall.checked;
    }
  }
 function loadThreadFollow(t_id,b_id){
	var targetImg =eval("document.all.followImg" + t_id);
	var targetDiv =eval("document.all.follow" + t_id);
	
	if ("object"==typeof(targetImg)){
		if (targetDiv.style.display!='block'){
			targetDiv.style.display="block";
			
		}else{
			targetDiv.style.display="none";
		}
	}
}
</script>

<script LANGUAGE="JavaScript">

<!--

function checkMobile(){

	var sMobile = document.mobileform.mobile.value

	if(!(/^((\(\d{3}\))|(\d{3}\-))?13[0-9]\d{8}|15[0-9]\d{8}|18[0-9]\d{8}/.test(sMobile))) //(�������15,18�����Ŷ�)

	{

		alert("����������11λ�ֻ��Ż�����ȷ���ֻ���ǰ��λ");

		document.mobileform.mobile.focus();

		return false;

	}
	//window.open('', 'mobilewindow' target="_blank")
}



//-->

</script>
<SCRIPT LANGUAGE="JavaScript">
<!--
function checkIP()
{
	var ipArray,ip,j;
	ip = document.ipform.ip.value;
	
	if(/[A-Za-z_-]/.test(ip)){
		if(!/^([\w-]+\.)+((com)|(net)|(org)|(gov\.cn)|(info)|(cc)|(com\.cn)|(net\.cn)|(org\.cn)|(name)|(biz)|(tv)|(cn))$/.test(ip)){
			alert("������ȷ������");
			document.ipform.ip.focus();
			return false;
		}
	}
	else{
		ipArray = ip.split(".");
		j = ipArray.length
		if(j!=4)
		{
			alert("������ȷ��IP");
			document.ipform.ip.focus();
			return false;
		}

		for(var i=0;i<4;i++)
		{
			if(ipArray[i].length==0 || ipArray[i]>255)
			{
				alert("������ȷ��IP");
				document.ipform.ip.focus();
				return false;
			}
		}
	}
}
//-->
</SCRIPT>
<%
set rs=server.createobject("adodb.recordset")
sql = "select * from DNJY_info where id="&cstr(id)
rs.open sql,conn,1,3
rs.update
%>
<html>
<head>
<title><%=titleinfo%>-<%=title%></title>
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name="keywords" content="<%=keywords%>">
<meta name="description" content="<%=description%>">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<style type="text/css">
<!--
body {
	background-image: url(images/back.gif);
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
<link href="../css/css.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.STYLE11 {
	color: #FFFFFF;
	font-weight: bold;
}
-->
</style>
</head>
<BODY>

<body>
<table width="500" border="0" align="center" cellpadding="0" cellspacing="0">

  <tr>
   
<%
set rs=server.createobject("adodb.recordset")
sql = "select * from DNJY_info where id="&cstr(id)
rs.open sql,conn,1,3

rs("llcs")=rs("llcs")+1
rs.update
%>
    <td width="500" valign="top" bgcolor="#FFFFFF"><table width="542" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
      <tr bgcolor="#E6E6E6">
        <td height="4"></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF"><table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
          <tr>
            <td height="5"></td>
            </tr>
          <tr>
            <td height="28" bgcolor="#52bed9"><table width="100%" border="0" cellspacing="0" cellpadding="5">
              <tr>
                <td><span class="STYLE11"><%=rs("biaoti")%> [<%=datevalue(rs("fbsj"))%>]</span></td>
              </tr>
            </table></td>
            </tr>
          <tr>
            <td><table width="100%" border="0" cellpadding="4" cellspacing="0" bgcolor="#FFFFFF">
              <tr>
                <td width="15%" bgcolor="#F9FCFD"><div align="center">��Ϣ���</div></td>
                <td width="256"><%
	belongtype="<a href=""../list.asp?t="&rs("type_oneid")&","&rs("type_twoid")&","&rs("type_threeid")&"&"&C_ID&""">"&rs("type_one")&"</a>"
	IF rs("type_two")<>"" and not isnull(rs("type_two")) Then belongtype=belongtype&" / <a href=""../list.asp?t="&rs("type_oneid")&","&rs("type_twoid")&","&rs("type_threeid")&"&"&C_ID&""">"&rs("type_two")&"</a>"
	IF rs("type_three")<>"" and not isnull(rs("type_three")) Then belongtype=belongtype&" / <a href=""../list.asp?t="&rs("type_oneid")&","&rs("type_twoid")&","&rs("type_threeid")&"&"&C_ID&""">"&rs("type_three")&"</a>"
	response.write belongtype%></td>
                <td width="220" rowspan="6"><div align="center">
                <%if rs("c")=0 then%>
               <img src="../images/nophoto.gif" alt="��ͼƬ" width="220" height="180" border="0">
               <%else 
                if rs("c")=1 and not rs("tupian")="0" then%><p align="center">
            <a target="_blank" title="����Ŵ�&gt;&gt;&gt;" href="../UploadFiles/infopic/<%=rs("tupian")%>">
            <img src="../UploadFiles/infopic/<%=rs("tupian")%>" alt="����Ŵ�" width="220" height="180" border="0"><br>[����Ŵ�]</a><%end if%><%end if%></div></td>
              </tr>
              <tr>
                <td bgcolor="#F9FCFD"><div align="center">��Ϣ���ͣ�</div></td>
                <td><%=rs("leixing")%></td>
                </tr>
              <tr>
                <td bgcolor="#F9FCFD"><div align="center">�г��۸�</div></td>
                <td><font color="#FF0000"><%=check_jiage(rs("scjiage"))%></font></td>
                </tr>
              <tr>
                <td bgcolor="#F9FCFD"><div align="center">��վ�۸�</div></td>
                <td><font color="#FF0000"><%=check_jiage(rs("jiage"))%></font></td>
                </tr>
              <tr>
                <td bgcolor="#F9FCFD"><div align="center">���׵�����</div></td>
                <td>
<%	belongcity="<a href=""../index.asp?c="&rs("city_oneid")&","&rs("city_twoid")&","&rs("city_threeid")&""">"&rs("city_one")&"</a>"
    IF rs("city_two")<>"" and not isnull(rs("city_two")) Then belongcity=belongcity&" / <a href=""../index.asp?c="&rs("city_oneid")&","&rs("city_twoid")&","&rs("city_threeid")&""">"&rs("city_two")&"</a>"
	IF rs("city_three")<>"" and not isnull(rs("city_three")) Then belongcity=belongcity&" / <a href=""../index.asp?c="&rs("city_oneid")&","&rs("city_twoid")&","&rs("city_threeid")&""">"&rs("city_three")&"</a>"
	response.write belongcity
%></td>
                </tr>
             <tr>
                <td bgcolor="#F9FCFD"><div align="center">��ԱID�ţ�</div></td>
                <td ><%	If IsNull(rs("username")) then%>��ע���Ա<%else%><%=rs("username")%> <a href=../preview.asp?username=<%=rs("username")%>&id=<%=id%>>����������Ϣ</a><%End if%>
				</td>				
                </tr>
              <tr>
                <td bgcolor="#F9FCFD"><div align="center">�� ϵ �ˣ�</div></td>
                <td><b><%=rs("name")%></b></td>
                </tr>
              <tr>
                <td bgcolor="#F9FCFD"><div align="center">�̶��绰��</div></td>
                <td colspan="2"><%=rs("dianhua")%></td>
            </td>
                </tr>
              <tr>
                 <form action="http://www.ip138.com:8080/search.asp" method="get" onSubmit="return checkMobile();" target="_blank" name="mobileform"><td bgcolor="#F9FCFD"><div align="center">�ƶ��绰��</div></td>
               <td><input type="hidden" name="action" value="mobile"><input type="text" name="mobile" size="11" value="<%=rs("CompPhone")%>" style="border: 1px solid #FFFFFF">
			  [<INPUT type=submit value=�鿴�绰 name=Submit style="border:1px solid #FFFFFF; background-color:#FFFFFF">]
            </td></form>
                </tr>

              <tr>
                <td bgcolor="#F9FCFD"><div align="center">IP�����أ�</div></td>
                <FORM METHOD=POST ACTION="http://www.ip138.com/ips.asp" name="ipform" target="_blank"><td colspan="2"><input type="text" name="ip" size="15" value="<%=rs("ip")%>" style="border: 1px solid #FFFFFF">[<input type="submit" value="�鿴IP" style="border:1px solid #FFFFFF; background-color:#FFFFFF"><INPUT TYPE="hidden" name="action" value="2">]</td></FORM></tr>
              <tr>
                <td bgcolor="#F9FCFD"><div align="center">Q Q ���룺</div></td>
                <td colspan="2"><%=rs("qq")%><a target=blank href=tencent://message/?Uin=<%=rs("qq")%>&Site=<%=web%>&Menu=yes><img border='0' SRC=http://wpa.qq.com/pa?p=1:<%=rs("qq")%>:7 alt=''��������'></a></td>
                </tr>
              <tr>
                <td bgcolor="#F9FCFD"><div align="center">E - mail��</div></td>
                <td colspan="2"><%=rs("email")%></td>
                </tr>
              <tr>
                <td bgcolor="#F9FCFD"><div align="center">��ϵ��ַ��</div></td>
                <td colspan="2"><%=rs("dizhi")%></td>
                </tr>
              <tr>
                <td bgcolor="#F9FCFD"><div align="center">����ظ���</div></td>
                <td colspan="2"><%=rs("llcs")%>����� / <%=rs("hfcs")%>���ظ�</td>
                </tr>
              <tr>
                <td bgcolor="#F9FCFD"><div align="center">����ʱ�䣺</div></td>
                <td colspan="2"><%dim sj
sj=DateDiff("d",now(),""&rs("dqsj")&"")
if sj>0 then
response.write "<font color=""#FF0000"">ʣ��"&sj&"</b>��</font>"

elseif sj=0 then
response.write "<font color=""#FF0000"">���յ���</font>"

elseif sj<0 then
response.write "<font color=""#FF0000"">�Ѿ�ʧЧ</font>"
end if%>&nbsp;[<%dim zt
zt=rs("zt")
if zt=1 then
response.write "<font color=""#ff0000"">�Ѿ���ɽ���</font>"
elseif zt=0 then
response.write "<font color=""#0066FF"">��δ��ɽ���</font>"

end if%>]</td>
                </tr>
              <tr>
                <td valign="top" bgcolor="#F9FCFD"><div align="center">��ϸ˵����</div></td>
                <td colspan="2" style="word-break:break-all"><%=HTMLDecode(rs("memo"))%></td>
                </tr>
              <tr>
                <td height="3" colspan="3" bgcolor="#FFFFFF"></td>
                </tr>
              <tr>
                <td height="2" colspan="3" bgcolor="#CCCCCC"></td>
              </tr>

            </table></td>
            </tr>

        </table></td>
      </tr>
    </table></td>
    <td width="4" bgcolor="#E6E6E6"></td>
  </tr>
</table>
</body>
</html>

<%
Call CloseRs(rs)
Call CloseDB(conn)
%>