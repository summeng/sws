<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file="top.asp"-->
<%dim zuo1,zuo2,zuo3,info1,info2,info3,info4,info5,info6,info7,info8,info9,xxtp,biaoti,leixin,biaotixs,fbsj,scjiage,jiage,sdmba,usernameid,namea,dianhua,mobile,userqq,email,dizhi,youbian,memo,hfcs,dqsj,sj
%>
<!--
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������Ұ�Ȩ����
'�ٷ��ͷ����� http://www.ip126.com
'����֧����̳ http://www.ip126.com/bbs
'QQ:530051328 mail:bdunion@126.com
'=====================================--->

<script language="JavaScript">
//�����ظ�����
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
	if(!(/^13[0-9]\d{4,8}$/.test(sMobile))){
		alert("����������11λ�ֻ��Ż�����ȷ���ֻ���ǰ��λ");
		document.mobileform.mobile.focus();
		return false;
	}
	window.open('', 'mobilewindow', 'height=197,width=350,status=yes,toolbar=no,menubar=no,location=no')
}
//-->
</script>


<html>
<head>
<title><%=biaoti%>-<%=title%></title>
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name="keywords" content="<%=keywords%>">
<meta name="description" content="<%=description%>">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script type="text/javascript" src="Include/copy.js"></script>
</head>
<body>
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="214" height="100%" valign="top"><!--#include file=left.asp--></td>
	<td width="5">&nbsp;</td>��
<%id=trim(request("id"))
set rs=server.createobject("adodb.recordset")
sql = "select * from DNJY_info where id="&id&" and username='"&request.cookies("dick")("username")&"'"
rs.open sql,conn,1,3
if rs.eof Then
Response.Write "<CENTER><p>û�з���������Ϣ</CENTER>"
response.end
end If
CT_ID=CT_ID
username=rs("username")
leixin=rs("leixing")
biaoti=rs("biaoti")
fbsj=rs("fbsj")
userip=rs("ip")
memo=rs("memo")
hfcs=rs("hfcs")
dianhua=rs("dianhua")
mobile=rs("CompPhone")
userqq=rs("qq")
email=rs("email")
dizhi=rs("dizhi")
youbian=rs("youbian")
CT_ID=CT_ID
username=rs("username")
leixin=rs("leixing")
biaoti=rs("biaoti")
fbsj=rs("fbsj")
userip=rs("ip")
memo=rs("memo")
hfcs=rs("hfcs")
dianhua=rs("dianhua")
mobile=rs("CompPhone")
userqq=rs("qq")
email=rs("email")
dizhi=rs("dizhi")
youbian=rs("youbian")
dqsj=rs("dqsj")
Readinfo=rs("Readinfo")
scjiage=check_jiage(rs("scjiage"))
jiage=check_jiage(rs("jiage"))
if rs("tj")>0 Then
biaotixs="<font color=""#FF0000""><img border=""0"" src=""images/jian.gif"" alt=""�Ƽ���Ϣ""></font>"
end If
if rs("a")="0" Then
biaotixs=biaotixs&""&rs("biaoti")&"</b>"
Else
biaotixs=biaotixs&"<font color=""#"&rs("a")&"""><b>"&rs("biaoti")&"</b></font>"
end If
biaotixs=biaotixs&"</font>"
If rs("username")<>"" And rs("username")<>"�ο�" then
dim rs6,sql6,vip,sdkg,sdname,mpkg,mpname,mpfg
Set rs6= Server.CreateObject("ADODB.Recordset")
sql6="select vip,name,sdkg,sdname,mpkg,mpname,mpfg from [DNJY_user] where username='"&rs("username")&"'"
rs6.open sql6,conn,1,1
vip=rs6("vip")
sdkg=rs6("sdkg")
sdname=rs6("sdname")
mpkg=rs6("mpkg")
mpname=rs6("mpname")
mpfg=rs6("mpfg")
rs6.close
set rs6=Nothing
end If

if sdkg=1 And sdname<>"" Then
sdmba="<a href=""user/com.asp?com="&rs("username")&""" target=""_blank""><img src=""img/dian.gif"" title=""�˻�Ա�ѿ��е���"" ></a>"
end if
if mpkg=1 And mpname<>"" then
sdmba=sdmba+" <a href=""company.asp?username="&rs("username")&""" target=""_blank""><img src=""img/qy.gif"" title=""�˻�Ա�ѿ�����ҵ��Ƭ"" ></a>"
end if
if sdkg=0 And  mpkg=0 Then
sdmba="�޵��̺���ҵ��Ƭ��δ�����"
end if



If IsNull(rs("username")) Then
usernameid="��ע���Ա"
Else
usernameid=""&rs("username")&" <a href=preview.asp?username="&rs("username")&"&id="&id&">����������Ϣ</a>"
End if

if vip=1 Then
namea="<img src=""images/vip.gif"" alt=""��֤VIP�û�"" width=""13"" height=""13"" border=""0"">&nbsp;"
end If
namea=namea&""&rs("name")&""
if rs("c")=0 then
xxtp="<script language=javascript src=""adjs_path(""ads/js"",""info4"","&c1&"_"&c2&"_"&c3&")""></script>"
else 
if rs("c")=1 and not rs("tupian")="0" then
xxtp="<a target=""_blank"" title=""����Ŵ�&gt;&gt;&gt;"" href=""UploadFiles/infopic/"&rs("tupian")&"""><img src=""UploadFiles/infopic/"&rs("tupian")&""" alt=""����Ŵ�"" width=""250"" height=""200"" border=""0""><br><center>[��ϢͼƬ ����Ŵ�]</a>"
end if
end if


	belongtype="<a href=""list.asp?t="&rs("type_oneid")&"&"&C_ID&""">"&rs("type_one")&"</a>"
	IF rs("type_two")<>"" and not isnull(rs("type_two")) Then belongtype=belongtype&" / <a href=""list.asp?t="&rs("type_oneid")&","&rs("type_twoid")&"&"&C_ID&""">"&rs("type_two")&"</a>"
	IF rs("type_three")<>"" and not isnull(rs("type_three")) Then belongtype=belongtype&" / <a href=""list.asp?t="&rs("type_oneid")&","&rs("type_twoid")&","&rs("type_threeid")&"&"&C_ID&""">"&rs("type_three")&"</a>"	


    belongcity="<a href=""index.asp?c="&rs("city_oneid")&""">"&rs("city_one")&"</a>"
    IF rs("city_two")<>"" and not isnull(rs("city_two")) Then belongcity=belongcity&" / <a href=""index.asp?c="&rs("city_oneid")&","&rs("city_twoid")&""">"&rs("city_two")&"</a>"
	IF rs("city_three")<>"" and not isnull(rs("city_three")) Then belongcity=belongcity&" / <a href=""index.asp?c="&rs("city_oneid")&","&rs("city_twoid")&","&rs("city_threeid")&""">"&rs("city_three")&"</a>"%>
    <td width="760" valign="top" bgcolor="#FFFFFF">
<table align="center" bgcolor="#ffffff" class="ty1"><tr><td ><script language=javascript src="<%=adjs_path("ads/js","info1",c1&"_"&c2&"_"&c3)%>"></script></td></tr></table>
      <table width="760" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
 <table width="760" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="ty1">
      <tr>
        <td bgcolor="#FFFFFF" ><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
          <tr>
            <td height="5"></td>
            </tr>
          <tr>
            <td height="28" ><table width="100%" border="0" cellspacing="0" cellpadding="5">
 <tr>
<td height="38" colspan="4" align="center" style="border-top: 1px solid #C0C0C0; padding-left: 0; padding-right: 0; padding-top: 1px; padding-bottom: 1px">
<font size="3"><b>

<%
response.write "["&leixin&"] "&biaotixs&""%>
</td>
</tr>
<tr><td height="27" colspan="4" style="border-bottom: 1px solid #C0C0C0; padding: 1px">
<p align="center"><b>����ʱ�䣺</b><%=fbsj%> <%if rs("dqsj")<>"" Then
sj=DateDiff("d",now(),""&rs("dqsj")&"")
if sj>0 then
response.Write "<font color=#FF0000> ʣ��"&sj&"</b>��</font>"
elseif sj=0 then
response.Write "<font color=""#414141"">���յ���</font>"
elseif sj<0 then
response.Write "<font color=""#FF0000""> �Ѿ�����</font>"
end If
end If%>��<b>���/�ظ���</b><%response.Write rs("llcs")%>��/<%response.Write rs("hfcs")%>��  ������¼��<SCRIPT src="user/xinfodt.asp?action=Bad&id=<%=id%>"></SCRIPT>�� <a href="Bad_info_list.asp?infoid=<%=cstr(id)%>&biaoti=<%=biaoti%>" target=blank>��</a> [<a title=ƾɾ���������ɾ������Ϣ href="#" ONCLICK="window.open('../user_delxx.asp?id=<%=id%>&key=del','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=300,height=200,left=300,top=100')">ɾ��</a>]</td></tr>
            </table></td>
            </tr></table>

              <table width="750" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="260" valign="top"> <table width="250" height="161" border="0" cellpadding="0" cellspacing="2" bgcolor="#E4E4E4">
                      <tr> 
                        <td height="250" align="center" bgcolor="#FFFFFF"><%=xxtp%></td>
                      </tr>
                    </table></td>
                  <td valign="top"><table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#E4E4E4">
                      <tr bgcolor="#FFFFFF"> 
                        <td width="50%">&nbsp;��Ϣ���<%=belongtype%></td>
                        <td>&nbsp;�� ϵ �ˣ�<%=namea%></td>
                      </tr>
                      <tr bgcolor="#FFFFFF"> 
                        <td>&nbsp;�������ͣ�[<%=leixin%>]</td>
                        <td>&nbsp;��ԱID�ţ�<%=usernameid%></td>
                      </tr>
                      <tr bgcolor="#FFFFFF"> 
                        <td>&nbsp;����״̬��<SCRIPT src="user/xinfodt.asp?action=zt&id=<%=id%>"></SCRIPT></td>
                        <td>&nbsp;�̶��绰��<%=dianhua%></td>
						
                      </tr>
                      <tr bgcolor="#FFFFFF"> 
                        <td>&nbsp;�г��۸�<%=scjiage%>/��վ�۸�<%=jiage%></td>
                         <form action="http://www.ip138.com:8080/search.asp" method="get" onSubmit="return checkMobile();" target="mobilewindow" name="mobileform"><td>&nbsp;�ƶ��绰��<input type="hidden" name="action" value="mobile">
			  <input type="text" name="mobile" size="16" value="<%=mobile%>" style="border: 1px solid #FFFFFF">
			  <INPUT type="submit" value="����" name="Submit">
                        </td></form>
                      </tr>
                      <tr bgcolor="#FFFFFF"> 
                        <td>&nbsp;���׵�����<%=belongcity%></td>
                        <td>&nbsp;�����ʼ���<%=email%> <a href="mailto:<%=email%>?subject=����<%=title%>��������������Ϣ��������ϵ">���Ÿ���</a></td>
                      </tr>
                      <tr bgcolor="#FFFFFF"> 
                        <td>&nbsp;�̼ҵ���<%=sdmba%></td>
<td>&nbsp;Q Q ���룺<%=userqq%><a target=blank href=http://wpa.qq.com/msgrd?V=1&Uin=<%=userqq%>&Site=<%=web%>&Menu=yes><img border='0' SRC=http://wpa.qq.com/pa?p=1:<%=userqq%>:7 alt=''��������'></a></td>
                      </tr>
                      <tr bgcolor="#FFFFFF"> 
                        <form method="get" action="http://www.ip138.com/ips1388.asp" name="ipform" onsubmit="return checkIP();" target="_blank">
	<td>&nbsp;I P ��ַ��<input type="text" name="ip" size="15" value="<%=userip%>" style="border: 1px solid #FFFFFF"> <input type="submit" value="����"><input type="hidden" name="action" value="2"></td></td></form>
                        <td>&nbsp;�������룺<%=youbian%></td>
                      </tr>
                      <tr bgcolor="#FFFFFF"> 
                        <td height="32" colspan="2">&nbsp;ͨѶ��ַ��<%=dizhi%></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td colspan="2" height="7"></td>
                </tr>
              </table>
 			  <table width="750" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td valign="top" bgcolor="#FFFFFF" class="a14">
				  <table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#E4E4E4">
                      <tr bgcolor="#FFFFFF"> 
                        <td width="50%" bgcolor="#FCFCFC"><div style="margin:9px; font-size:14px; letter-spacing:1px; line-height:140%"><%=HTMLDecode(memo)%></div></td></tr></table></td>
                </tr>
                <tr> 
                  <td height="7"></td>
                </tr>
              </table>
<table width="760" border="1" cellpadding="0" cellspacing="0" align="center">
        <tr>
          <td  style="border-style:none; padding:1px;" valign="top"><script language=javascript src="<%=adjs_path("ads/js","tail",c1&"_"&c2&"_"&c3)%>"></script></td>
</tr></table>             
<%Call CloseRs(rs)%>
<table width="760" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
              <tr>
                <td colspan="3" bgcolor="#FFFFFF" align="center">
                <INPUT class="inputb" TYPE=button VALUE="��վ�ڶ��Ÿ��û�" ONCLICK="window.open('messh.asp?name=<%=username%>','Sample')">
                <INPUT class="inputb" TYPE=button VALUE="�����ҵ��ղ�" ONCLICK="window.open('user/collection.asp?id=<%=Id%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=170,height=80,left=300,top=100')">
                <input type="button" name="Submit" onClick='copyToClipBoard()' value="�����Ƽ�������">
				<INPUT class="inputb" TYPE=button VALUE="�ٱ�����Ϣ" ONCLICK="window.open('user/Bad_info.asp?id=<%=Id%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=500,height=600,left=300,top=100')">
				</td>

              </tr>
              <tr>
                <td height="10" colspan="3" align="center">������ժ��| <a href="javascript:window.open('http://shuqian.qq.com/post?from=3&title='+encodeURIComponent(document.title)+'&uri='+encodeURIComponent(document.location.href)+'&jumpback=2&noui=1','favit','width=930,height=470,left=50,top=50,toolbar=no,menubar=no,location=no,scrollbars=yes,status=yes,resizable=yes');void(0)" style="background:url('http://shuqian.qq.com/img/add.gif') no-repeat 0px 0px;height:23px;width:60px;padding:2px 2px 0px 20px;font-size:12px;">QQ��ǩ</a> | <A title=����ǿ��������ղؼУ�һ���Ӳ����Ϳ�������ʵ�ֱ�������ļ�ֵ����������Ŀ��� href="javascript:d=document;t=d.selection?(d.selection.type!='None'?d.selection.createRange().text:''):(d.getSelection?d.getSelection():'');void(keyit=window.open('http://www.365key.com/storeit.aspx?t='+escape(d.title)+'&u='+escape(d.location.href)+'&c='+escape(t),'keyit','scrollbars=no,width=475,height=575,left=75,top=20,status=no,resizable=yes'));keyit.focus();"><FONT color=darkorchid>365K<FONT color=#57c200>e</FONT>y</FONT></A> |  <A title="�ղص�POCO ��ժhttp://share.poco.cn" href="javascript:d=document;t=d.selection?(d.selection.type!='None'?d.selection.createRange().text:''):(d.getSelection?d.getSelection():'');void(vivi=window.open('http://my.poco.cn/fav/storeIt.php?t='+escape(d.title)+'&u='+escape(d.location.href)+'&c='+escape(t),'_blank','scrollbars=no,width=475,height=575,left=75,top=20,status=no,resizable=yes'));"><FONT color=#4cb509>POCO��ժ</FONT></A> |<A style="FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #000000; TEXT-DECORATION: none" href="javascript:window.open('http://cang.baidu.com/do/add?it='+encodeURIComponent(document.title.substring(0,76))+'&iu='+encodeURIComponent(location.href)+'&fr=ien#nw=1','_blank','scrollbars=no,width=600,height=450,left=75,top=20,status=no,resizable=yes'); void 0"><SPAN style="PADDING-RIGHT: 5px; PADDING-LEFT: 5px; FONT-SIZE: 12px; PADDING-BOTTOM: 0px; MARGIN-LEFT: 10px; CURSOR: pointer; PADDING-TOP: 5px">�ٶ���ժ</A> |</SPAN></td>
              </tr>

              <tr>
                <td height="10" colspan="3" bgcolor="#FFFFFF"></td>
              </tr>
<script language="javascript">
<!--
//��֤������ȷ��
function checkemail(mdz){
var str=mdz;
//��JavaScript�У�������ʽֻ��ʹ��"/"��ͷ�ͽ���������ʹ��˫����
var Expression=/\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*/;
var objExp=new RegExp(Expression);
if(objExp.test(str)==true){
return true;
}else{
return false;
}
}

function form_onsubmit() {
    if((!checkemail(form.mdz.value)) && (document.form.mdz.value!=0))
	{
    alert("������Email��ַ����ȷ������������!");
    document.form.mdz.focus();
    return false;
    }
if (document.form.neirong.value==0)
	{
	  alert("����ظ�����")
	  document.form.neirong.focus()
	  return false
	 }
}
// --></script>


            </table>

  <!---��Ϣ��ϸҳ�ײ�һ��濪ʼ---->
<script language=javascript src="<%=adjs_path("ads/js","info9",c1&"_"&c2&"_"&c3)%>"></script>
  <!---��Ϣ��ϸҳ�ײ�һ������---->



  </td>
            </tr>

        </table></td>
      </tr>
    </table></td>
    <td width="4" bgcolor="#E6E6E6"></td>
  </tr>
</table>
<table align="center" bgcolor="#ffffff">
<tr><td>
<script language=javascript src="<%=adjs_path("ads/js","tc",c1&"_"&c2&"_"&c3)%>"></script>
</td></tr>
</table>
</body>
</html>

<!--#include file="footer.asp" -->