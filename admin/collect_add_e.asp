<!--#include file="../conn.asp"-->
<!--#include file=cookies.asp-->
<!--#include file=../config.asp-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
response.buffer=true
%>
<%Call OpenConn
Dim ItemID
Dim RsItem,SqlItem,FoundErr,ErrMsg
Dim UrlTest,TsString,ToString,CsString,CoString
Dim DateType,DsString,DoString,DateTime
Dim StopDateType,SDsString,SDoString,StopDateTime,StopDate
Dim AuthorType,AsString,AoString,AuthorStr,TelType,TelsString,TeloString,TelStr,sjType,sjsString,sjoString,sjStr,QQType,QsString,QoString,QQStr,dzType,dzsString,dzoString,dzStr,ybType,ybsString,yboString,ybStr,jgType,jgsString,jgoString,jgStr
Dim CopyFromType,FsString,FoString,CopyFromStr,leixingType,LxsString,LxoString,leixingStr,mailType,mailsString,mailoString,mailStr


Dim NewsUrl,NewsCode
Dim ConTent,Author,CopyFrom,jg,sj,dz,yb,telc,mail,qqc,leixingc
Dim strChannelDir
Dim cj_strInstallDir

cj_strInstallDir=trim(request.ServerVariables("SCRIPT_NAME"))
cj_strInstallDir=left(cj_strInstallDir,instrrev(lcase(cj_strInstallDir),"/")-1)
cj_strInstallDir=left(cj_strInstallDir,instrrev(lcase(cj_strInstallDir),"/"))
strChannelDir="Test"

ItemID=strint(Request.Form("ItemID"))           '��ĿID
UrlTest=Trim(Request.Form("UrlTest"))           'һ����Ϣ
TsString=Request.Form("TsString")               '�����ǩ��ʼ
ToString=Request.Form("ToString")               '�����ǩ����
CsString=Request.Form("CsString")               '���ı�ǩ��ʼ
CoString=Request.Form("CoString")               '���ı�ǩ����

DateType=strint(Request.Form("DateType"))       '����ʱ������ 0.��ǰʱ�� 1.���ñ�ǩ
DsString=Request.Form("DsString")               '����ʱ���ǩ��ʼ
DoString=Request.Form("DoString")               '����ʱ���ǩ����

StopDateType=strint(Request.Form("StopDateType")) '����ʱ������ 0.��ǰʱ��+ 1.���ñ�ǩ
SDsString=Request.Form("SDsString")             '����ʱ���ǩ��ʼ
SDoString=Request.Form("SDoString")             '����ʱ���ǩ����
StopDate=strint(request.form("StopDate"))       '����ʱ��+

jgType=strint(Request.Form("jgType"))   '���׼۸�
jgsString=Request.Form("jgsString")     
jgoString=Request.Form("jgoString")     
jgStr=Trim(Request.Form("jgStr"))  

AuthorType=strint(Request.Form("AuthorType"))   '������������ 0.�������� 1.���ñ�ǩ 2.ָ������
AsString=Request.Form("AsString")               '���߱�ǩ��ʼ
AoString=Request.Form("AoString")               '���߱�ǩ����
AuthorStr=Trim(Request.Form("AuthorStr"))       'ָ������

TelType=strint(Request.Form("TelType"))   '�绰�������� 0.�������� 1.���ñ�ǩ 2.ָ���绰
TelsString=Request.Form("TelsString")               '�绰��ǩ��ʼ
TeloString=Request.Form("TeloString")               '�绰��ǩ����
TelStr=Trim(Request.Form("telStr"))       'ָ���绰

sjType=strint(Request.Form("sjType"))   '�ֻ��������� 0.�������� 1.���ñ�ǩ 2.ָ������
sjsString=Request.Form("sjsString")               '�ֻ���ǩ��ʼ
sjoString=Request.Form("sjoString")               '�ֻ���ǩ����
sjStr=Trim(Request.Form("sjStr"))       'ָ���ֻ�

mailType=strint(Request.Form("mailType"))   '������������ 0.�������� 1.���ñ�ǩ 2.ָ������
mailsString=Request.Form("mailsString")               '�����ǩ��ʼ
mailoString=Request.Form("mailoString")               '�����ǩ����
mailStr=Trim(Request.Form("mailStr"))       'ָ������

QQType=strint(Request.Form("QQType"))   'QQ�������� 0.�������� 1.���ñ�ǩ 2.ָ��QQ
QsString=Request.Form("QsString")               'QQ��ǩ��ʼ
QoString=Request.Form("QoString")               'QQ��ǩ����
QQStr=Trim(Request.Form("QQStr"))       'ָ��QQ

dzType=strint(Request.Form("dzType"))   '��ַ
dzsString=Request.Form("dzsString")     
dzoString=Request.Form("dzoString")     
dzStr=Trim(Request.Form("dzStr"))       

ybType=strint(Request.Form("ybType"))   '�ʱ�
ybsString=Request.Form("ybsString")      
yboString=Request.Form("yboString")      
ybStr=Trim(Request.Form("ybStr"))   

CopyFromType=strint(Request.Form("CopyFromType"))'��Ϣ��Դ�������� 0.�������� 1.���ñ�ǩ 2.ָ����Դ 3.���ɼ�ҳ
FsString=Request.Form("FsString")               '��Ϣ��Դ��ǩ��ʼ
FoString=Request.Form("FoString")               '��Ϣ��Դ��ǩ����
CopyFromStr=Trim(Request.Form("CopyFromStr"))   'ָ����Դ

leixingType=strint(Request.Form("leixingType"))'��Ϣ������������ 0.�������� 1.���ñ�ǩ 2.ָ������
LxsString=Request.Form("LxsString")               '��Ϣ���ͱ�ǩ��ʼ
LxoString=Request.Form("LxoString")               '��Ϣ���ͱ�ǩ����
leixingStr=Trim(Request.Form("leixingStr"))   'ָ������

If ItemID=0 Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>�������������Ч���ӽ���</li>"
End If

If UrlTest="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>�����������ݴ���ʱ��������</li>"
Else
   NewsUrl=UrlTest
End If

If TsString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>���⿪ʼ��ǲ���Ϊ��</li>"
End If
If ToString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>���������ǲ���Ϊ��</li>" 
End If
If CsString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>���Ŀ�ʼ��ǲ���Ϊ��</li>"
End If
If CoString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>���Ľ�����ǲ���Ϊ��</li>" 
End If


If DateType=1 Then
  If DsString="" or DoString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>�뽫����ʱ��Ŀ�ʼ/���������д����</li>" 
  End If
End If

If StopDateType=1 Then
  If SDsString="" or SDoString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>�뽫��ʱ��ʱ��Ŀ�ʼ/���������д����</li>" 
  End If
End If

If jgType=1 Then
  If jgsString="" or jgoString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>�뽫���׼۸�Ŀ�ʼ/���������д����</li>" 
  End If
ElseIf jgType=2 Then
  If jgStr="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>��ָ�����׼۸�</li>" 
  End If
End If
If IsNumeric(jgStr)=False And jgType=1 Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>���׼۸�ֻ�������֣�</li>"
End If
      

If AuthorType=1 Then
  If AsString="" or AoString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>�뽫���ߵĿ�ʼ/���������д����</li>" 
  End If
ElseIf AuthorType=2 Then
  If AuthorStr="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>��ָ������</li>" 
  End If
End If

If TelType=1 Then
  If TelsString="" or TeloString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>�뽫�̶��绰�Ŀ�ʼ/���������д����</li>" 
  End If
ElseIf TelType=2 Then
  If TelStr="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>��ָ���̶��绰</li>" 
  End If
End If

If sjType=1 Then
  If sjsString="" or sjoString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>�뽫�ֻ�����Ŀ�ʼ/���������д����</li>" 
  End If
ElseIf sjType=2 Then
  If sjStr="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>��ָ���ֻ�����</li>" 
  End If
End If

If mailType=1 Then
  If mailsString="" or mailoString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>�뽫����Ŀ�ʼ/���������д����</li>" 
  End If
ElseIf mailType=2 Then
  If mailStr="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>��ָ������</li>" 
  End If
End If

If QQType=1 Then
  If QsString="" or QoString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>�뽫QQ�Ŀ�ʼ/���������д����</li>" 
  End If
ElseIf QQType=2 Then
  If QQStr="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>��ָ��QQ</li>" 
  End If
End If

If dzType=1 Then
  If dzsString="" or dzoString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>�뽫ͨѶ��ַ�Ŀ�ʼ/���������д����</li>" 
  End If
ElseIf dzType=2 Then
  If dzStr="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>��ָ��ͨѶ��ַ</li>" 
  End If
End If

If ybType=1 Then
  If ybsString="" or yboString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>�뽫�ʱ�Ŀ�ʼ/���������д����</li>" 
  End If
ElseIf ybType=2 Then
  If ybStr="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>��ָ���ʱ�</li>" 
  End If
End If

If CopyFromType=1 Then
  If FsString="" or FoString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>�뽫��Դ�Ŀ�ʼ/���������д������</li>" 
  End If
ElseIf CopyFromType=2 Then
  If CopyFromStr="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>��ָ����Դ</li>" 
  End If
End If

If leixingType=1 Then
  If LxsString="" or LxoString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>�뽫���͵Ŀ�ʼ/���������д������</li>" 
  End If
ElseIf leixingType=2 Then
  If leixingStr="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>��ָ������</li>" 
  End If
End If

'�����ݿ�



If FoundErr<>True Then
   SqlItem="Select * from DNJY_xx_Item Where ItemID=" & ItemID
   Set RsItem=server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,Conn,2,3

   RsItem("TsString")=TsString
   RsItem("ToString")=ToString
   RsItem("CsString")=CsString
   RsItem("CoString")=CoString

   RsItem("DateType")=DateType
   If DateType=1 Then
      RsItem("DsString")=DsString
      RsItem("DoString")=DoString
   End If

   RsItem("StopDateType")=StopDateType
   If StopDateType=0 Then
	  RsItem("StopDate")=StopDate
   ElseIf StopDateType=1 Then
	  RsItem("SDsString")=SDsString
	  RsItem("SDoString")=SDoString
   End If

   RsItem("jgType")=jgType
   If jgType=1 Then
      RsItem("jgsString")=jgsString
      RsItem("jgoString")=jgoString
   ElseIf jgType=2 Then
      RsItem("jgStr")=jgStr
   End If

   RsItem("AuthorType")=AuthorType
   If AuthorType=1 Then
      RsItem("AsString")=AsString
      RsItem("AoString")=AoString
   ElseIf AuthorType=2 Then
      RsItem("AuthorStr")=AuthorStr
   End If

   RsItem("TelType")=TelType
   If TelType=1 Then
      RsItem("TelsString")=TelsString
      RsItem("TeloString")=TeloString
   ElseIf TelType=2 Then
      RsItem("TelStr")=TelStr
   End If

   RsItem("sjType")=sjType
   If sjType=1 Then
      RsItem("sjsString")=sjsString
      RsItem("sjoString")=sjoString
   ElseIf sjType=2 Then
      RsItem("sjStr")=sjStr
   End If

   RsItem("mailType")=mailType
   If mailType=1 Then
      RsItem("mailsString")=mailsString
      RsItem("mailoString")=mailoString
   ElseIf mailType=2 Then
      RsItem("mailStr")=mailStr
   End If

   RsItem("QQType")=QQType
   If QQType=1 Then
      RsItem("QsString")=QsString
      RsItem("QoString")=QoString
   ElseIf QQType=2 Then
      RsItem("QQStr")=QQStr
   End If

   RsItem("dzType")=dzType
   If dzType=1 Then
      RsItem("dzsString")=dzsString
      RsItem("dzoString")=dzoString
   ElseIf dzType=2 Then
      RsItem("dzStr")=dzStr
   End If

   RsItem("ybType")=ybType
   If ybType=1 Then
      RsItem("ybsString")=ybsString
      RsItem("yboString")=yboString
   ElseIf ybType=2 Then
      RsItem("ybStr")=ybStr
   End If

   RsItem("CopyFromType")=CopyFromType
   If CopyFromType=1 Then
      RsItem("FsString")=FsString
      RsItem("FoString")=FoString
   ElseIf CopyFromType=2 Then
      RsItem("CopyFromStr")=CopyFromStr
   End If

   RsItem("leixingType")=leixingType
   If leixingType=1 Then
      RsItem("LxsString")=LxsString
      RsItem("LxoString")=LxoString
   ElseIf leixingType=2 Then
      RsItem("leixingStr")=leixingStr
   End If

   RsItem.UpDate
   RsItem.Close
   Set RsItem=Nothing
End If
'�ر����ݿ�����
conn.close:Set conn=nothing

If FoundErr<>True Then
   NewsCode=GetHttpPage(NewsUrl)
   If NewsCode<>"$False$" Then
      Title=GetBody(NewsCode,TsString,ToString,False,False)
      Content=GetBody(NewsCode,CsString,CoString,False,False)
      If Title="$False$" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>�ڽ�ȡ�����ʱ��������" & NewsUrl & "</li>"
      Else
         Title=FpHtmlEnCode(Title)
      End If

      If Content="$False$" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>�ڽ�ȡ���ĵ�ʱ��������" & NewsUrl & "</li>"
      End If
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>�ڻ�ȡԴ��ʱ��������"& NewsUrl &"</li>"
   End If 
End If

If FoundErr<>True Then
      If DateType=0 then
         DateTime=date()
      Else
         DateTime=GetBody(NewsCode,DsString,DoString,False,False)
         DateTime=FpHtmlEncode(DateTime)
         DateTime=Trim(Replace(DateTime,"&nbsp;"," "))
         If IsDate(DateTime)=True Then
            DateTime=CDate(DateTime)
         Else
            DateTime=date()
         End If
      End If
	  DateTime=FormatDateTime(DateTime,2)

      If StopDateType=0 then
         StopDateTime=dateadd("d",StopDate,date())
      Else
         StopDateTime=GetBody(NewsCode,SDsString,SDoString,False,False)
         StopDateTime=FpHtmlEncode(StopDateTime)
         StopDateTime=Trim(Replace(StopDateTime,"&nbsp;"," "))
         If IsDate(StopDateTime)=True Then
            StopDateTime=CDate(StopDateTime)
         Else
            StopDateTime=dateadd("d",31,date())
         End If
      End If
	  StopDateTime=FormatDateTime(StopDateTime,2)

      If jgType=1 Then
         jg=GetBody(NewsCode,jgsString,jgoString,False,False)
      ElseIf jgType=2 Then
         jg=jgStr
      End If
      If jg="$False$" Or IsNull(jg) Then
         jg=0
      Else
      If IsInteger(jg)=True then
         jg=CLng(jg)
      Else
      jg=0
      End If
      End If

      If AuthorType=1 Then
         Author=GetBody(NewsCode,AsString,AoString,False,False)
      ElseIf AuthorType=2 Then
         Author=AuthorStr
      End If
      If Author="$False$" Or IsNull(Author) Then
         Author="����"
      Else
         Author=FpHtmlEnCode(Author)
      End If

      If TelType=1 Then
         Telc=GetBody(NewsCode,TelsString,TeloString,False,False)
      ElseIf TelType=2 Then
         Telc=TelStr
      End If
      If Telc="$False$" Or IsNull(Telc) Then
         Telc="δ��"
      Else
         Telc=FpHtmlEnCode(Telc)
      End If

      If sjType=1 Then
         sj=GetBody(NewsCode,sjsString,sjoString,False,False)
      ElseIf sjType=2 Then
         sj=sjStr
      End If
      If sj="$False$" Or IsNull(sj) Then
         sj="δ��"
      Else
         sj=FpHtmlEnCode(sj)
      End If

      If mailType=1 Then
         mail=GetBody(NewsCode,mailsString,mailoString,False,False)
      ElseIf mailType=2 Then
         mail=mailStr
      End If
      If mail="$False$" Or IsNull(mail) Then
         mail="δ��"
      Else
         mail=FpHtmlEnCode(mail)
      End If

      If QQType=1 Then
         QQc=GetBody(NewsCode,qsString,qoString,False,False)
      ElseIf qqType=2 Then
         qqc=qqStr
      End If
      If qqc="$False$" Or IsNull(qqc) Then
         qqc="δ��"
      Else
         qqc=FpHtmlEnCode(qqc)
      End If

      If dzType=1 Then
         dz=GetBody(NewsCode,dzsString,dzoString,False,False)
      ElseIf dzType=2 Then
         dz=dzStr
      End If
      If dz="$False$" Or IsNull(dz) Then
         dz="δ��"
      Else
         dz=FpHtmlEnCode(dz)
      End If

      If ybType=1 Then
         yb=GetBody(NewsCode,ybsString,yboString,False,False)
      ElseIf ybType=2 Then
         yb=ybStr
      End If
      If yb="$False$" Or IsNull(yb) Then
         yb="δ��"
      Else
         yb=FpHtmlEnCode(yb)
      End If
	  
      If CopyFromType=1 Then
         CopyFrom=GetBody(NewsCode,FsString,FoString,False,False)
      ElseIf CopyFromType=2 Then
         CopyFrom=CopyFromStr
      ElseIf CopyFromType=3 Then
         CopyFrom=NewsUrl
      End If
      If CopyFrom="$False$" Or IsNull(CopyFrom) Then
         CopyFrom="����"
      Else
         CopyFrom=FpHtmlEnCode(CopyFrom)
      End If

      If leixingType=1 Then
         leixingc=GetBody(NewsCode,LxsString,LxoString,False,False)
      ElseIf leixingType=2 Then
         leixingc=leixingStr
      End If
      If leixingc="$False$" Or IsNull(leixingc) Then
         leixingc="δ��"
      Else
         leixingc=FpHtmlEnCode(leixingc)
      End If

End If

If FoundErr<>True Then
   Content=ReplaceSaveRemoteFile(Content,cj_strInstallDir,strChannelDir,False,NewsUrl)
End If

If FoundErr=True Then
   Call WriteErrMsg(ErrMsg)
Else
   Call Main()
End if

%>
<%Sub Main()%>
<HTML>
<HEAD>
<TITLE>���ݲɼ�ϵͳ</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
</HEAD>
<BODY leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<TABLE width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <TR> 
    <TD height="22" colspan="2" align="center" class="topbg"><strong>��&nbsp;&nbsp;��&nbsp;&nbsp;ϵ&nbsp;&nbsp;ͳ&nbsp;&nbsp;��&nbsp;&nbsp;Ŀ&nbsp;&nbsp;��&nbsp;&nbsp;��</strong></TD>
  </TR>
  <TR class="tdbg"> 
    <TD width="65" height="30"><STRONG>��������</STRONG></TD>
    <TD height="30"><A href="collect_add_a.asp">�����Ŀ</A> >> <A href="collect_editor_a.asp?ItemID=<%=ItemID%>">��������</A> >> <A href="collect_editor_b.asp?ItemID=<%=ItemID%>">�б�����</A> >> <A href="collect_editor_c.asp?ItemID=<%=ItemID%>">��������</A> >> <A href="collect_editor_d.asp?ItemID=<%=ItemID%>">��������</A> >> <FONT color=red>��������</FONT> >> �������� >> ���</TD>         
  </TR>         
</TABLE> 
<BR>        
<TABLE width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <TR> 
      <TD height="22" colspan="2" class="title"> <DIV align="center"><STRONG>�� �� �� �� Ŀ--�� �� �� ��</STRONG></DIV></TD>
    </TR>
</TABLE>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
  <TR align="center" class="tdbg"> 
    <TD height="40" colspan="2" valign="bottom"><SPAN lang="zh-cn"><font size="5"><%=Title%></font></SPAN></TD>
  </TR>
  <TR align="center" class="tdbg">
    <TD colspan="2">
        ��Դ��<%=CopyFrom%>&nbsp;&nbsp;���ߣ�<%=Author%>&nbsp;&nbsp;����ʱ�䣺<%=DateTime%>&nbsp;&nbsp;����ʱ�䣺<%=StopDateTime%>
    </TD>
  </TR>
  <TR class="tdbg">
    <TD colspan="2">
      <TABLE width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">
        <TR>
          <TD height="200" valign="top"><P><SPAN lang="zh-cn"><%=Content%></SPAN></P>
          </TD>
        </TR>
      </TABLE>
    </TD>
  </TR>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
  <tr class="tdbg">
    <td colspan="2">
      <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">
        <tr>
          <td height="200" valign="top"><b>�ɼ��������������Ŀ��</b><p>
		  <p>��Ϣ���ͣ�<%=leixingc%></p><p>���׼۸�<%=jg%></p><p>�̻���<%=Telc%></p><p>�ֻ���<%=sj%></p><p>QQ��<%=QQc%></p><p>���䣺<%=mail%></p><p>ͨѶ��ַ��<%=dz%></p><p>�ʱࣺ<%=yb%></p>
</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<FORM method="post" action="collect_attribute.asp" name="form1">
<TABLE width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <TR class="tdbg"> 
      <TD colspan="2" align="center" class="tdbg">
        <INPUT name="ItemID" type="hidden" value="<%=ItemID%>">
        <INPUT name="button1" type="button" id="Cancel" value=" ��&nbsp;һ&nbsp;�� " onClick="window.location.href='javascript:history.go(-1)'" style="cursor: hand;background-color: #cccccc;">
        &nbsp; 
        <INPUT  type="submit" name="Submit" value="  ��&nbsp;һ&nbsp;�� " style="cursor: hand;background-color: #cccccc;"></TD>
    </TR>
</TABLE>
</FORM>       
</BODY>         
</HTML>
<%End Sub%>


