<!--#include file="../conn.asp"-->
<!--#include file=cookies.asp-->
<!--#include file=../config.asp-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file="inc/Function.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
response.buffer=true	
%>

<%Call OpenConn
Dim ItemID
Dim RsItem,SqlItem,FoundErr,ErrMsg
Dim ListStr,LsString,LoString
Dim  ListPaingType,LPsString,LPoString,ListPaingStr1,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3,HsString,HoString,HttpUrlType,HttpUrlStr
Dim  ListUrl,ListCode,NewsArrayCode,NewsArray,UrlTest,NewsCode
dim Testi
ItemID=Trim(Request("ItemID"))
HsString=Request.Form("HsString")
HoString=Request.Form("HoString")
HttpUrlType=Trim(Request.Form("HttpUrlType"))
HttpUrlStr=Trim(Request.Form("HttpUrlStr"))


If ItemID="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>"
Else
   ItemID=Clng(ItemID)
End If
If HsString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>链接开始标记不能为空</li>"
End If
If HoString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>链接结束标记不能为空</li>" 
End If
If HttpUrlType="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>请选择链接处理类型</li>"
Else
   HttpUrlType=Clng(HttpUrlType)
   If HttpUrlType=1 Then
      If HttpUrlStr="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>绝对链接字符设置不能为空</li>"
      Else
            If  Len(HttpUrlStr)<15  Then
               FoundErr=True
               ErrMsg=ErrMsg & "<br><li>绝对链接字符设置不正确(15个字符以上)</li>"
            End  If      
      End If
   End If
End If


If FoundErr<>True Then
   SqlItem="Select ListStr,LsString,LoString,ListPaingType,LPsString,LPoString,ListPaingStr1,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3,HsString,HoString,HttpUrlType,HttpUrlStr from DNJY_xx_Item Where ItemID=" & ItemID
   Set RsItem=server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,Conn,2,3

   RsItem("HsString")=HsString
   RsItem("HoString")=HoString
   RsItem("HttpUrlType")=HttpUrlType
   If HttpUrlType=1 Then
      RsItem("HttpUrlStr")=HttpUrlStr
   End If
   ListStr=RsItem("ListStr")
   LsString=RsItem("LsString")
   LoString=RsItem("LoString")
   ListPaingType=RsItem("ListPaingType")
   LPsString=RsItem("LPsString")
   ListPaingStr1=RsItem("ListPaingStr1")
   ListPaingStr2=RsItem("ListPaingStr2")
   ListPaingID1=RsItem("ListPaingID1")
   ListPaingID2=RsItem("ListPaingID2")
   ListPaingStr3=RsItem("ListPaingStr3")

   RsItem.UpDate
   RsItem.Close
   Set RsItem=Nothing
   
   Select  Case  ListPaingType
   Case  0,1
         ListUrl=ListStr
   Case  2
         ListUrl=Replace(ListPaingStr2,"{$ID}",CStr(ListPaingID1))
   Case  3
         If  Instr(ListPaingStr3,"|")>0  Then
         ListUrl=Left(ListPaingStr3,Instr(ListPaingStr3,"|")-1)
      Else
         ListUrl=ListPaingStr3
      End  If
   End  Select

   ListCode=GetHttpPage(ListUrl)
   If ListCode<>"$False$" Then
      ListCode=GetBody(ListCode,LsString,LoString,False,False)
      If ListCode<>"$False$" Then 
         NewsArrayCode=GetArray(ListCode,HsString,HoString,False,False)
         If NewsArrayCode<>"$False$" Then
               NewsArray=Split(NewsArrayCode,"$Array$")
               For Testi=0 To Ubound(NewsArray)
                  If HttpUrlType=1 Then
                     NewsArray(Testi)=Replace(HttpUrlStr,"{$ID}",NewsArray(Testi))
                  Else
                     NewsArray(Testi)=DefiniteUrl(NewsArray(Testi),ListUrl)
                  End If
               Next
               UrlTest=NewsArray(0)
               NewsCode=GetHttpPage(UrlTest)
         Else
            FoundErr=True
            ErrMsg=ErrMsg & "<br><li>在获取链接列表时出错。</li>"
         End If   
      Else
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>在截取列表时发生错误。</li>"
      End If
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>在获取:" & ListUrl & "网页源码时发生错误。</li>"
   End If
End If

If FoundErr=True Then
   Call WriteErrMsg(ErrMsg)
Else
   Call Main()
End If 
'关闭数据库链接
conn.close:Set conn=nothing
%>
<%Sub Main()%>
<html>
<head>
<title>采集系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr> 
    <td height="22" colspan="2" align="center" class="topbg"><strong>采&nbsp;&nbsp;集&nbsp;&nbsp;系&nbsp;&nbsp;统&nbsp;&nbsp;项&nbsp;&nbsp;目&nbsp;&nbsp;管&nbsp;&nbsp;理</td>
  </tr>
  <tr class="tdbg"> 
    <td width="65" height="30"><strong>管理导航：</strong></td>
    <td height="30"><a href="collect_add_a.asp">添加项目</a> >> <a href="collect_editor_a.asp?ItemID=<%=ItemID%>">基本设置</a> >> <a href="collect_editor_b.asp?ItemID=<%=ItemID%>">列表设置</a> >> <a href="collect_editor_c.asp?ItemID=<%=ItemID%>">链接设置</a> >> <font color=red>正文设置</font> >> 采样测试 >> 属性设置 >> 完成</td>         
  </tr>         
</table> 
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >        
    <tr> 
      <td height="22" colspan="2" class="title"> <div align="center"><strong>添 加 新 
		项 目--列 表 新 闻 链 接 测 试</strong></div></td>
    </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <tr> 
      <td height="22" colspan="2" class="tdbg">以下是分析后所得到的新闻绝对链接地址，请查看是否正确。<br>
<%
For Testi=0 To Ubound(NewsArray)
   Response.Write "<a href='" & NewsArray(Testi) & "' target=_blank>" & NewsArray(Testi) & "</a><br>"
Next
%>
<br>
下一步将抽取第一条新闻进行测试，在填写以下标记时尽量不要使用第一条新闻。
      </td>
    </tr>
</table>
<form method="post" action="collect_add_e.asp" name="form1">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <tr> 
      <td height="22" colspan="2" class="title"> <div align="center"><strong>添 加 新 
		项 目--正 文 设 置</strong></div>      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg" ><strong>标题开始标记：</strong><p>　</p><p>　</p>
      <strong>标题结束标记：</strong></td>
      <td class="tdbg" width="75%">
      <textarea name="TsString" cols="49" rows="7"></textarea><br>
      <textarea name="ToString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg" ><strong>正文开始标记：</strong><p>　</p><p>　</p>
      <strong>正文结束标记：</strong></td>
      <td class="tdbg" width="75%">
      <textarea name="CsString" cols="49" rows="7"></textarea><br>
      <textarea name="CoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg" ><b>发布时间设置：</b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="DateType" checked onClick="Date1.style.display='none'">当前时间&nbsp;
	    <input type="radio" value="1" name="DateType" onClick="Date1.style.display=''">设置标签&nbsp;    （不作设置时取当前时间）</tr>
    <tr class="tdbg" id="Date1" style="display:none"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>发布时间开始标记：</font></strong><p>　</p>
		<p>　</p>
      <strong><font color=blue>发布时间结束标记：</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="DsString" cols="49" rows="7"></textarea><br>
      <textarea name="DoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg" ><b>到期时间设置：</b></td>
      <td class="tdbg" width="75%">
      	<input name="StopDateType" type="radio" onClick="Date2.style.display='none';Date3.style.display=''" value="0" checked>&nbsp;当前时间+&nbsp;↓
   	<input type="radio" value="1" name="StopDateType" onClick="Date2.style.display='';Date3.style.display='none'">设置标签&nbsp;</tr>
	
	<tr class="tdbg" id="Date2" style="display:none"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>到期时间开始标记：</font></strong>
        <p>　</p>
		<p>　</p>
      <strong><font color=blue>到期时间结束标记：</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="SDsString" cols="49" rows="7"></textarea><br>
      <textarea name="SDoString" cols="49" rows="7"></textarea></td>
    </tr>
	
    <tr class="tdbg" id="Date3"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>请选择到期时间:</font></strong></td>
      <td class="tdbg" width="75%">
      	<SELECT size="1" name="StopDate" id="StopDate">
          <OPTION value="7" selected>一星期</OPTION>
          <OPTION value="15">半个月</OPTION>
          <OPTION value="30">一个月</OPTION>          
          <OPTION value="90">三个月</OPTION>
          <OPTION value="180">半年</OPTION>
          <OPTION value="365">一年</OPTION>
          <OPTION value="1100">三年</OPTION>
        </SELECT></td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg" ><b><span lang="en-us">交易价格设置：</span></b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="jgType" checked onClick="jg1.style.display='none';jg2.style.display='none'">不作设置&nbsp;
		<input type="radio" value="1" name="jgType" onClick="jg1.style.display='';jg2.style.display='none'">设置标签&nbsp;
		<input type="radio" value="2" name="jgType" onClick="jg1.style.display='none';jg2.style.display=''">
		指定交易价格</td>
    </tr>
    <tr class="tdbg" id="jg1" style="display:none"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>交易价格开始标记：</font></strong>
        <p>　</p>
		<p>　</p>
      <strong><font color=blue>交易价格结束标记：</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="jgsString" cols="49" rows="7"></textarea><br>
      <textarea name="jgoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr class="tdbg" id="jg2" style="display:none"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>请指定交易价格：</font></strong></td>
      <td class="tdbg" width="75%">
      <input name="jgStr" type="text" id="jgStr" value="" onKeyUp="if(isNaN(value)){alert('交易价格只允许输入数字');value='';}"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))">  （必须为数字！）</td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg" ><b><span lang="en-us">作&nbsp; &nbsp;者&nbsp;&nbsp;设&nbsp; &nbsp;</span>置：</b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="AuthorType" checked onClick="Author1.style.display='none';Author2.style.display='none'">不作设置&nbsp;
		<input type="radio" value="1" name="AuthorType" onClick="Author1.style.display='';Author2.style.display='none'">设置标签&nbsp;
		<input type="radio" value="2" name="AuthorType" onClick="Author1.style.display='none';Author2.style.display=''">指定作者（请设置，否则采集后信息编辑时也要填）</td>
    </tr>
    <tr class="tdbg" id="Author1" style="display:none"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>作者开始标记：</font></strong><p>　</p>
		<p>　</p>
      <strong><font color=blue>作者结束标记：</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="AsString" cols="49" rows="7"></textarea><br>
      <textarea name="AoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr class="tdbg" id="Author2" style="display:none"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>请指定作者：</font></strong></td>
      <td class="tdbg" width="75%">
      <input name="AuthorStr" type="text" id="AuthorStr" value="">      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg" ><b><span lang="en-us">固定电话设置：</span></b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="TelType" checked onClick="Tel1.style.display='none';Tel2.style.display='none'">不作设置&nbsp;
		<input type="radio" value="1" name="TelType" onClick="Tel1.style.display='';Tel2.style.display='none'">设置标签&nbsp;
		<input type="radio" value="2" name="TelType" onClick="Tel1.style.display='none';Tel2.style.display=''">
		指定电话</td>
    </tr>
    <tr class="tdbg" id="Tel1" style="display:none"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>电话开始标记：</font></strong>
        <p>　</p>
		<p>　</p>
      <strong><font color=blue>电话结束标记：</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="TelsString" cols="49" rows="7"></textarea><br>
      <textarea name="TeloString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr class="tdbg" id="Tel2" style="display:none"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>请指定电话：</font></strong></td>
      <td class="tdbg" width="75%">
      <input name="TelStr" type="text" id="TelStr" value="">      </td>
    </tr>

    <tr class="tdbg"> 
      <td width="20%" class="tdbg" ><b><span lang="en-us">手机号码设置：</span></b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="sjType" checked onClick="sj1.style.display='none';sj2.style.display='none'">不作设置&nbsp;
		<input type="radio" value="1" name="sjType" onClick="sj1.style.display='';sj2.style.display='none'">设置标签&nbsp;
		<input type="radio" value="2" name="sjType" onClick="sj1.style.display='none';sj2.style.display=''">
		指定手机号码</td>
    </tr>

    <tr class="tdbg" id="sj1" style="display:none"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>手机号码开始标记：</font></strong>
        <p>　</p>
		<p>　</p>
      <strong><font color=blue>手机号码结束标记：</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="sjsString" cols="49" rows="7"></textarea><br>
      <textarea name="sjoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr class="tdbg" id="sj2" style="display:none"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>请指定手机号码：</font></strong></td>
      <td class="tdbg" width="75%">
      <input name="sjStr" type="text" id="sjStr" value="">      </td>
    </tr>

	 <tr class="tdbg">
<td width="20%" class="tdbg" ><b><span lang="en-us">电&nbsp;&nbsp;子&nbsp;&nbsp;信&nbsp;&nbsp;箱：</span></b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="mailType" checked onClick="mail1.style.display='none';mail2.style.display='none'">不作设置&nbsp;
		<input type="radio" value="1" name="mailType" onClick="mail1.style.display='';mail2.style.display='none'">设置标签&nbsp;
		<input type="radio" value="2" name="mailType" onClick="mail1.style.display='none';mail2.style.display=''">指定信箱（请设置，否则采集后信息编辑时也要填）</td>
    </tr>
    <tr class="tdbg" id="mail1" style="display:none"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>信箱开始标记：</font></strong>
        <p>　</p>
		<p>　</p>
      <strong><font color=blue>信箱结束标记：</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="mailsString" cols="49" rows="7"></textarea><br>
      <textarea name="mailoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr class="tdbg" id="mail2" style="display:none"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>请指定信箱：</font></strong></td>
      <td class="tdbg" width="75%">
      <input name="mailStr" type="text" id="mailStr" value=""></td>
    </tr>

   <tr class="tdbg"> 
      <td width="20%" class="tdbg" ><b><span lang="en-us">QQ&nbsp;&nbsp;设&nbsp;&nbsp;置：</span></b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="QQType" checked onClick="QQ1.style.display='none';QQ2.style.display='none'">不作设置&nbsp;
		<input type="radio" value="1" name="QQType" onClick="QQ1.style.display='';QQ2.style.display='none'">设置标签&nbsp;
		<input type="radio" value="2" name="QQType" onClick="QQ1.style.display='none';QQ2.style.display=''">
		指定QQ</td>
    </tr>
    <tr class="tdbg" id="QQ1" style="display:none"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>QQ开始标记：</font></strong>
        <p>　</p>
		<p>　</p>
      <strong><font color=blue>QQ结束标记：</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="QsString" cols="49" rows="7"></textarea><br>
      <textarea name="QoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr class="tdbg" id="QQ2" style="display:none"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>请指定QQ：</font></strong></td>
      <td class="tdbg" width="75%">
      <input name="QQStr" type="text" id="QQStr" value="">      </td>
    </tr>

    <tr class="tdbg"> 
      <td width="20%" class="tdbg" ><b><span lang="en-us">通讯地址设置：</span></b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="dzType" checked onClick="dz1.style.display='none';dz2.style.display='none'">不作设置&nbsp;
		<input type="radio" value="1" name="dzType" onClick="dz1.style.display='';dz2.style.display='none'">设置标签&nbsp;
		<input type="radio" value="2" name="dzType" onClick="dz1.style.display='none';dz2.style.display=''">
		指定通讯地址</td>
    </tr>

    <tr class="tdbg" id="dz1" style="display:none"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>通讯地址开始标记：</font></strong>
        <p>　</p>
		<p>　</p>
      <strong><font color=blue>通讯地址结束标记：</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="dzsString" cols="49" rows="7"></textarea><br>
      <textarea name="dzoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr class="tdbg" id="dz2" style="display:none"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>请指定通讯地址：</font></strong></td>
      <td class="tdbg" width="75%">
      <input name="dzStr" type="text" id="dzStr" value="">      </td>
    </tr>

    <tr class="tdbg"> 
      <td width="20%" class="tdbg" ><b><span lang="en-us">邮政编码设置：</span></b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="ybType" checked onClick="yb1.style.display='none';yb2.style.display='none'">不作设置&nbsp;
		<input type="radio" value="1" name="ybType" onClick="yb1.style.display='';yb2.style.display='none'">设置标签&nbsp;
		<input type="radio" value="2" name="ybType" onClick="yb1.style.display='none';yb2.style.display=''">
		指定邮编</td>
    </tr>

    <tr class="tdbg" id="yb1" style="display:none"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>邮编开始标记：</font></strong>
        <p>　</p>
		<p>　</p>
      <strong><font color=blue>邮编结束标记：</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="ybsString" cols="49" rows="7"></textarea><br>
      <textarea name="yboString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr class="tdbg" id="yb2" style="display:none"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>请指定邮编：</font></strong></td>
      <td class="tdbg" width="75%">
      <input name="ybStr" type="text" id="ybStr" value="">      </td>
    </tr>

    <tr class="tdbg"> 
      <td width="20%" class="tdbg" ><b>来&nbsp; 源&nbsp;
		设&nbsp; 置：</b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="CopyFromType" checked onClick="CopyFrom1.style.display='none';CopyFrom2.style.display='none'">不作设置&nbsp;
		<input type="radio" value="1" name="CopyFromType" onClick="CopyFrom1.style.display='';CopyFrom2.style.display='none'">设置标签&nbsp;
		<input type="radio" value="2" name="CopyFromType" onClick="CopyFrom1.style.display='none';CopyFrom2.style.display=''">指定来源
		<input type="radio" value="3" name="CopyFromType" onClick="CopyFrom1.style.display='none';CopyFrom2.style.display='none'">
		被采集页</td>
    </tr>
    <tr class="tdbg" id="CopyFrom1" style="display:none"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>来源开始标记：</font></strong><p>　</p>
		<p>　</p>
      <strong><font color=blue>来源结束标记：</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="FsString" cols="49" rows="7"></textarea><br>
      <textarea name="FoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr class="tdbg" id="CopyFrom2" style="display:none"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>请指定来源：</font></strong></td>
      <td class="tdbg" width="75%">
      <input name="CopyFromStr" type="text" id="CopyFromStr" value="">      </td>
    </tr>
   <tr class="tdbg"> 
      <td width="20%" class="tdbg" ><b><span lang="en-us">信 息 类 型：</span></b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="leixingType" checked onClick="leixing1.style.display='none';leixing2.style.display='none'">不作设置&nbsp;
		<input type="radio" value="1" name="leixingType" onClick="leixing1.style.display='';leixing2.style.display='none'">设置标签&nbsp;
		<input type="radio" value="2" name="leixingType" onClick="leixing1.style.display='none';leixing2.style.display=''">
		指定类型</td>
    </tr>
    <tr class="tdbg" id="leixing1" style="display:none"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>类型开始标记：</font></strong>
        <p>　</p>
		<p>　</p>
      <strong><font color=blue>类型结束标记：</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="LxsString" cols="49" rows="7"></textarea><br>
      <textarea name="LxoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr class="tdbg" id="leixing2" style="display:none"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>请指定类型：</font></strong></td>
      <td class="tdbg" width="75%"><%Dim rslx,sqllx,leixing,x,le
set rslx=server.createobject("adodb.recordset")
sqllx = "select leixing from DNJY_config "
rslx.open sqllx,conn,1,1
leixing=split(rslx("leixing"),"|")
response.write "<select name=""leixingStr"" id=""leixingStr""><option value=>类型</option>"
for x=0 to ubound(leixing)
le=""
response.write "<option value="""&leixing(x)&""" "&le&">"&leixing(x)&"</option>"
next
response.write "</select>"
rslx.close
set rslx=nothing%></td>
    </tr>
    <tr class="tdbg"> 
      <td colspan="2" align="center" class="tdbg"><br>
          <input name="ItemID" type="hidden" value="<%=ItemID%>"> 
        <input  type="button" name="button1" value="上&nbsp;一&nbsp;步" onClick="window.location.href='javascript:history.go(-1)'"  style="cursor: hand;background-color: #cccccc;">
        &nbsp;&nbsp;&nbsp;&nbsp; 
        <input  type="submit" name="Submit" value="下&nbsp;一&nbsp;步" style="cursor: hand;background-color: #cccccc;"></td>
        <input type="hidden" name="UrlTest" id="UrlTest" value="<%=UrlTest%>">
    </tr>
</table>
</form>     
</body>         
</html>
<%End Sub%>
