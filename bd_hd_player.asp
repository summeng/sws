<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file="config.asp"-->
<!--#include file="Include/head.asp"-->
<!--#include file="Include/Function.asp"-->
<!--#include file="Include/label.asp"-->
<%'======================================
'�������ƣ����ţ���Ϣ��ͼ�Ļõ�Ƭ���Ų��
'�ص㣺����ʾ�����ʽͼƬ
'���÷�Χ�����칩����Ϣ��������ϵͳ�����������ݿ���ʽ���õ����ţ���Ϣ��ͼ�Ļõ���ʾ
'���л�����asp+access
'���÷�ʽ�������ļ��ŵ�WEB��Ŀ¼�����κ�ҳ��λ�ü��ϣ�<iframe marginwidth="0" marginheight="0" src="bd_hd_player.asp?MaxNum=10&PicWidth=280&PicHeight=250&Title=20&TextHeight=15" frameborder="0" width="280" height="250" scrolling="No"  align="center"></iframe>
'����˵����MaxNum -- ͼƬ��ʾ����
'          TextHeight -- �������ָ߶�
'����������PicWidth -- ͼƬ���
'����������PicHeight -- ͼƬ�߶�
'����������Title -- �����ַ�����
'
MaxNum=trim(request("MaxNum"))
PicWidth=trim(request("PicWidth"))
PicHeight=trim(request("PicHeight"))
Title=strint(request("Title"))
TextHeight=trim(request("TextHeight"))

'��������������Ƽ�(http://www.ip126.com)
'���ʹ�ã��뱣����Ȩ��Ϣ
'��ϵ���䣺bbfhq@sohu.com
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" >
<title>���ţ���Ϣ��ͼ�Ļõ���ʾ</title>
<style type="text/css">
body,th,td{font-size:12px}
a:link {
	text-decoration: none;color:#000000;
}
a:visited {
	text-decoration: none;color:#000000;
}
a:hover {
	text-decoration: none;color:#FF0000;
}
a:active {
	text-decoration: none;color:#FF0000;
}
</style>
</head>

<body>
<%Call OpenConn
set rsbdhd=server.createobject("ADODB.recordset")
sqlbdhd="select  top 100 * from [DNJY_info] where yz=1 and (tupian <>'0')"&ttcc&" order by fbsj "&DN_OrderType&""
rsbdhd.open sqlbdhd,conn,3,3
if rsbdhd.eof Then
Response.Write "<p><br>����"
IF c1>0 Then Response.Write conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
Response.Write "����ͼƬ��Ϣ!"
Call CloseRs(rsbdhd)
else%>

<SCRIPT LANGUAGE="JavaScript">
<!-- Hide
function killErrors() {
return true;
}
window.onerror=killErrors;
// -->
</SCRIPT>
<script language=JavaScript>
  var imgUrl=new Array();
  var imgLink=new Array();
  var ImgTitle=new Array();
  var adNum=0;
<%dim rsbdhd,sqlbdhd,MaxNum,TextHeight,PicWidth,PicHeight,AdID,AdUrl,AdTitle,AdPic,k


//��������Ϊ���ݿ���ã��ɸ��������ݿ������޸�
set rsbdhd=server.createobject("ADODB.recordset")
sqlbdhd="select top 100 * from [DNJY_info] where yz=1 and (tupian <>'0') order by fbsj "&DN_OrderType&""
rsbdhd.open sqlbdhd,conn,3,3

	k=rsbdhd.RecordCount
	If K>MaxNum Then k=MaxNum
	i=1
do while not rsbdhd.eof
	AdID = rsbdhd("id") //IDΪ��ϢID���ֶΣ��ɸ��������ݿ������޸�
	AdTitle = CutStr(rsbdhd("biaoti"),Title,"...") //bobaiΪ��Ϣ�����ֶΣ��ɸ��������ݿ������޸�
	AdPic = "UploadFiles/infopic/"& rsbdhd("tupian") //ͼƬ����·����UploadFiles/infopicĿ¼��tupian�ֶοɸ��������ݿ⼰��վĿ¼������޸�
	AdUrl = ""&x_path("",rsbdhd("id"),"",rsbdhd("Readinfo"))&"" //show.aspΪ�������ļ����ɸ�������վ�ļ�������޸�

	//���´��������Ҹģ����������ʾ������===========================================================
	Response.Write"imgUrl["&i&"]=""" & AdPic & """;" & vbCrLf
	Response.Write"ImgTitle["&i&"]=""" & AdTitle & """;" & vbCrLf
	Response.Write"imgLink["&i&"]=""" & AdUrl & """;" & vbCrLf
	If i>=k Then Exit Do	
	i=i+1	
	rsbdhd.movenext
    loop
Call CloseRs(rsbdhd)
%> 

function setTransition(){
  if (document.all){
     imgUrlrotator.filters.revealTrans.Transition=Math.floor(Math.random()*23);
     imgUrlrotator.filters.revealTrans.apply();
  }
}

function playTransition(){
  if (document.all)
     imgUrlrotator.filters.revealTrans.play()
}

function nextAd(){
  if(adNum<<%=k%>)adNum++;
     else adNum=1;
  setTransition();
  document.images.imgUrlrotator.src=imgUrl[adNum];
  document.all.tags("div")[0].innerText=ImgTitle[adNum]
  playTransition();
  theTimer=setTimeout("nextAd()", 8000);
}

function jump2url(){
  jumpUrl=imgLink[adNum];
  jumpTarget='_blank';
  if (jumpUrl != ''){
     if (jumpTarget != '')window.open(jumpUrl,jumpTarget);
     else location.href=jumpUrl;
  }
}
function displayStatusMsg() { 
  status=imgLink[adNum];
  document.returnValue = true;
}
window.onload=nextAd;/*���д��������Ǽ���IE7/8*/
//-->
</script>
<a onMouseOver="displayStatusMsg();return document.returnValue" href="javascript:jump2url()"><img style="FILTER: revealTrans(duration=2,transition=20)" height=<%=PicHeight%> src="javascript:nextAd()" width=<%=PicWidth%> border=0 name=imgUrlrotator>
<div style=" background-color:#FFFFFF;width:<%=PicWidth-1%>px; text-align:center;text-valign:middle;cursor:hand;line-height:<%=TextHeight%>px;"></div></a>
<%End If
Call CloseDB(conn)%>
</body>
</html>
