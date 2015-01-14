<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Dim UserArticle,IsURLRewrite
Const maxPagesize=2000				'���÷�ҳ����

Function InsertPageBreak(strText)
	Dim strPagebreak,s,ss
	Dim i,IsCount,c,iCount,strTemp,Temp_String,Temp_Array
	strPagebreak="[hiweb_break]"
	s=strText
	If Len(s)<maxPagesize Then
		InsertPageBreak=s
	End If
	s=Replace(s, strPagebreak, "")
	s=Replace(s, "&nbsp;", "<&nbsp;>")
	s=Replace(s, "&gt;", "<&gt;>")
	s=Replace(s, "&lt;", "<&lt;>")
	s=Replace(s, "&quot;", "<&quot;>")
	s=Replace(s, "&#39;", "<&#39;>")
	If s<>"" and maxPagesize<>0 and InStr(1,s,strPagebreak)=0 then
		IsCount=True
		Temp_String=""
		For i= 1 To Len(s)
			c=Mid(s,i,1)
			If c="<" Then
				IsCount=False
			ElseIf c=">" Then
				IsCount=True
			Else
				If IsCount=True Then
					If Abs(Asc(c))>255 Then
						iCount=iCount+2
					Else
						iCount=iCount+1
					End If
					If iCount>=maxPagesize And i<Len(s) Then
						strTemp=Left(s,i)
						If CheckPagination(strTemp,"table|a|b>|i>|strong|div|span") then
							Temp_String=Temp_String & Trim(CStr(i)) & "," 
							iCount=0
						End If
					End If
				End If
			End If	
		Next
		If Len(Temp_String)>1 Then Temp_String=Left(Temp_String,Len(Temp_String)-1)
		Temp_Array=Split(Temp_String,",")
		For i = UBound(Temp_Array) To LBound(Temp_Array) Step -1
			ss = Mid(s,Temp_Array(i)+1)
			If Len(ss) > 380 Then
				s=Left(s,Temp_Array(i)) & strPagebreak & ss
			Else
				s=Left(s,Temp_Array(i)) & ss
			End If
		Next
	End If
	s=Replace(s, "<&nbsp;>", "&nbsp;")
	s=Replace(s, "<&gt;>", "&gt;")
	s=Replace(s, "<&lt;>", "&lt;")
	s=Replace(s, "<&quot;>", "&quot;")
	s=Replace(s, "<&#39;>", "&#39;")
	InsertPageBreak=s
End Function

Function CheckPagination(strTemp,strFind)
	Dim i,n,m_ingBeginNum,m_intEndNum
	Dim m_strBegin,m_strEnd,FindArray
	strTemp=LCase(strTemp)
	strFind=LCase(strFind)
	If strTemp<>"" and strFind<>"" then
		FindArray=split(strFind,"|")
		For i = 0 to Ubound(FindArray)
			m_strBegin="<"&FindArray(i)
			m_strEnd  ="</"&FindArray(i)
			n=0
			do while instr(n+1,strTemp,m_strBegin)<>0
				n=instr(n+1,strTemp,m_strBegin)
				m_ingBeginNum=m_ingBeginNum+1
			Loop
			n=0
			do while instr(n+1,strTemp,m_strEnd)<>0
				n=instr(n+1,strTemp,m_strEnd)
				m_intEndNum=m_intEndNum+1
			Loop
			If m_intEndNum=m_ingBeginNum then
				CheckPagination=True
			Else
				CheckPagination=False
				Exit Function
			End If
		Next
	Else
		CheckPagination=False
	End If
End Function

Function ContentPagination(hiwebstr)
	Dim ContentLen, maxperpage, Paginate,CurrentPage,ArticleContent
	Dim arrContent, strContent, i
	Dim m_strFileUrl,m_strFileExt,ArticleID
	ArticleID=Request.QueryString("ID")
	strContent = InsertPageBreak(hiwebstr)
	ContentLen = Len(strContent)
	CurrentPage=Request.QueryString("Page")
	If CurrentPage="" Then CurrentPage=0
	If InStr(strContent, "[hiweb_break]") <= 0 Then
		ArticleContent = "<div id=""NewsContentLabel"" class=""NewsContent"">" & strContent & "</div><div id=""Message"" class=""Message""></div>"
	Else
		arrContent = Split(strContent, "[hiweb_break]")
		Paginate = UBound(arrContent) + 1
		If CurrentPage = 0 Then
			CurrentPage = 1
		Else
			CurrentPage = CLng(CurrentPage)
		End If
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > Paginate Then CurrentPage = Paginate
		strContent = "<div id=""NewsContentLabel"" class=""NewsContent"">"& arrContent(CurrentPage - 1)

		ArticleContent = ArticleContent & strContent
		If UserArticle = True Then
			ArticleContent = ArticleContent & "</p></div><div id=""Message"" class=""Message""></div><p align=""center""><b>"
		Else
			ArticleContent = ArticleContent & "</p></div><p align=""center""><b>"
		End If
		If IsURLRewrite Then
			m_strFileUrl = ArticleID & "_"
		Else
			m_strFileExt = ""
			m_strFileUrl = "?id=" & ArticleID & "&Page="
		End If
		If CurrentPage > 1 Then
			If IsURLRewrite And (CurrentPage-1) = 1 Then
				ArticleContent = ArticleContent & "<a href="""& ArticleID & m_strFileExt & """>��һҳ</a>&nbsp;&nbsp;"
			Else
				ArticleContent = ArticleContent & "<a href="""& m_strFileUrl & CurrentPage - 1 & m_strFileExt & """>��һҳ</a>&nbsp;&nbsp;"
			End If
		End If
		For i = 1 To Paginate
			If i = CurrentPage Then
				ArticleContent = ArticleContent & "<font color=""red"">[" & CStr(i) & "]</font>&nbsp;"
			Else
				If IsURLRewrite And i = 1 Then
					ArticleContent = ArticleContent & "<a href="""& ArticleID & m_strFileExt & """>[" & i & "]</a>&nbsp;"
				Else
					ArticleContent = ArticleContent & "<a href="""& m_strFileUrl & i & m_strFileExt & """>[" & i & "]</a>&nbsp;"
				End if
			End If
		Next
		If CurrentPage < Paginate Then
			ArticleContent = ArticleContent & "&nbsp;<a href="""& m_strFileUrl & CurrentPage + 1 & m_strFileExt & """>��һҳ</a>"
		End If
		ArticleContent = ArticleContent & "</b></p>"
	End If
	Response.Write(ArticleContent)
End Function
%>