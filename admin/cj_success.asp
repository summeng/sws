<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="inc/Article_Function.asp"-->
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
Dim SqlItem,RsItem,FoundErr,ErrMsg
Dim ItemID,ItemName
Dim Hits,UpDateType,UpDateTime,IncludePicYn,DefaultPicYn,OnTop,Hot
Dim Script_Iframe,Script_Object,Script_Script,Script_Div,Script_Class,Script_Span,Script_Img,Script_Font,Script_A,Script_Html,Script_Table,Script_Tr,Script_Td
Dim CollecListNum,CollecNewsNum,Passed,SaveFiles,CollecOrder,LinkUrlYn,cj_watermark
Dim tClass,tSpecial
FoundErr=False
ItemID=Request("ItemID")

If ItemID="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>�������������Ч���ӽ���</li>"
Else
   ItemID=Clng(ItemID)
End If

If FoundErr<>True Then
   Hits=strint(Request.Form("Hits"))              '����
   IncludePicYn=Trim(Request.Form("IncludePicYn"))'����ͼƬ
   DefaultPicYn=Trim(Request.Form("DefaultPicYn"))'ͼƬ�Ƽ�
   OnTop=Trim(Request.Form("OnTop"))'�̶���Ϣ
   Hot=Trim(Request.Form("Hot"))'�ȵ���Ϣ
   '��ǩ����
   Script_Iframe=Trim(Request.Form("Script_Iframe"))
   Script_Object=Trim(Request.Form("Script_Object"))
   Script_Script=Trim(Request.Form("Script_Script"))
   Script_Div=Trim(Request.Form("Script_Div"))
   Script_Class=Trim(Request.Form("Script_Class"))
   Script_Span=Trim(Request.Form("Script_Span"))
   Script_Img=Trim(Request.Form("Script_Img"))
   Script_Font=Trim(Request.Form("Script_Font"))
   Script_A=Trim(Request.Form("Script_A"))
   Script_Html=Trim(Request.Form("Script_Html"))
   Script_Table=Trim(Request.Form("Script_Table"))
   Script_Tr=Trim(Request.Form("Script_Tr"))
   Script_Td=Trim(Request.Form("Script_Td"))
   '��ǩ����
   CollecListNum=strint(Request.Form("CollecListNum"))'�б����
   CollecNewsNum=strint(Request.Form("CollecNewsNum"))'��Ϣ����
   Passed=Trim(Request.Form("Passed"))'���
   SaveFiles=Trim(Request.Form("SaveFiles"))'����ͼƬ
   CollecOrder=Trim(Request.Form("CollecOrder"))'����ɼ�
   LinkUrlYn=Trim(Request.Form("LinkUrlYn"))' URL��ת
   cj_watermark=Trim(Request.Form("cj_watermark"))'ˮӡ





End If
If FoundErr<>True Then
   SqlItem="Select * from DNJY_Article_Item Where ItemID=" & ItemID
   Set RsItem=server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,Conn,2,3
   If RsItem.Eof And RsItem.Bof Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>��������û���Ҹ���Ŀ</li>"
   Else
     ItemName =RsItem("ItemName")
      RsItem("Hits")=Hits
      RsItem("UpdateType")=UpdateType
      If UpdateType=2 Then
         RsItem("UpDateTime")=UpDateTime
      End If
      If IncludePicYn="yes" Then
         RsItem("IncludePicYn")=-1
      Else
         RsItem("IncludePicYn")=0
      End If
      If DefaultPicYn="yes" Then
         RsItem("DefaultPicYn")=-1
      Else
         RsItem("DefaultPicYn")=0
      End If
      If OnTop="yes" Then
         RsItem("OnTop")=-1
      Else
         RsItem("OnTop")=0
      End If
      If Hot="yes" Then
         RsItem("Hot")=-1
      Else
         RsItem("Hot")=0
      End If
	  
      If Script_Iframe="yes" Then
         RsItem("Script_Iframe")=-1
      Else
         RsItem("Script_Iframe")=0
      End If
      If Script_Object="yes" Then
         RsItem("Script_Object")=-1
      Else
         RsItem("Script_Object")=0
      End If
      If Script_Script="yes" Then
         RsItem("Script_Script")=-1
      Else
         RsItem("Script_Script")=0
      End If
      If Script_Div="yes" Then
         RsItem("Script_Div")=-1
      Else
         RsItem("Script_Div")=0
      End If
      If Script_Class="yes" Then
         RsItem("Script_Class")=-1
      Else
         RsItem("Script_Class")=0
      End If
      If Script_Span="yes" Then
         RsItem("Script_Span")=-1
      Else
         RsItem("Script_Span")=0
      End If
      If Script_Img="yes" Then
         RsItem("Script_Img")=-1
      Else
         RsItem("Script_Img")=0
      End If

      If Script_Font="yes" Then
         RsItem("Script_Font")=-1
      Else
         RsItem("Script_Font")=0
      End If
      If Script_A="yes" Then
         RsItem("Script_A")=-1
      Else
         RsItem("Script_A")=0
      End If

      If Script_Html="yes" Then
         RsItem("Script_Html")=-1
      Else
         RsItem("Script_Html")=0
      End If
      RsItem("CollecListNum")=CollecListNum
      RsItem("CollecNewsNum")=CollecNewsNum

      If Passed="yes" Then
         RsItem("Passed")=-1
      Else
         RsItem("Passed")=0
      End If
      If SaveFiles="yes" Then
         RsItem("SaveFiles")=-1
      Else
         RsItem("SaveFiles")=0
      End If
      If CollecOrder="yes" Then
         RsItem("CollecOrder")=-1
      Else
         RsItem("CollecOrder")=0
      End If
      If LinkUrlYn="yes" Then
         RsItem("LinkUrlYn")=-1
      Else
         RsItem("LinkUrlYn")=0
      End If
      RsItem("InputerType")=0

      RsItem("Flag")=True
      If Script_Table="yes" Then
         RsItem("Script_Table")=-1
      Else
         RsItem("Script_Table")=0
      End If
      If Script_Tr="yes" Then
         RsItem("Script_Tr")=-1
      Else
         RsItem("Script_Tr")=0
      End If
      If Script_Td="yes" Then
         RsItem("Script_Td")=-1
      Else
         RsItem("Script_Td")=0
      End If
      If cj_watermark="yes" Then
         RsItem("cj_watermark")=-1
      Else
         RsItem("cj_watermark")=0
      End If
      

   End If
   ErrMsg="<br><li>" & ItemName & " ��Ŀ��������ˣ�</li>"
   RsItem.UpDate
   RsItem.Close
   Set RsItem=Nothing 
End If

If FoundErr=True Then
   Call WriteErrMsg(ErrMsg)
Else
   Call WriteSucced(ErrMsg)
End If
'�ر����ݿ�����
Call CloseDB(conn)
%>
