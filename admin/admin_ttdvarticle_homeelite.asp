<% Response.charset="gb2312" %>
<!--#include file="../Config.asp"-->
<%
If Request("HomeEliteKey") <> HomeEliteKey Then Response.End
If Request("Action") = "create" Then
	Dim Width,Height,FontFamily,FontSize,PrintWords,FontColor
	Dim Jpeg,JpegLeft,JpegTop,JpegName
	JWidth		=	Cint(Request("homeEliteWidth"))
	JHeight		=	Cint(Request("homeEliteHeight"))
	FontFamily	=	Request("homeEliteFontFamily")
	FontSize	=	Cint(Request("homeEliteFontSize"))
	PrintWords	=	Trim(Request("biaoti"))
	FontColor	=	Request("homeEliteColor")
	
	PrintWords	=	" " & PrintWords & " "
	FontColor	=	"&H" & FontColor
	tempWidth	=	Int(GetStrLen(PrintWords)*1.024/2 * FontSize)
	JpegLeft	=	0
	JpegTop		=	0
	If tempWidth < JWidth Then JpegLeft = Int((JWidth - tempWidth)/2)
	If FontSize < JHeight Then JpegTop = Int((JHeight - FontSize)/2)
	
	Set Jpeg	=	Server.CreateObject("Persits.Jpeg")
	Jpeg.Open Server.MapPath("Images/" & "blank.jpg")
	If tempWidth < JWidth Then
		Jpeg.Width = JWidth
	Else
		Jpeg.Width = tempWidth
	End If
	Jpeg.Height =	JHeight
	Jpeg.Canvas.Font.Color	=	FontColor
	Jpeg.Canvas.Font.Family	=	FontFamily
	Jpeg.Canvas.Font.Size	=	FontSize
	Jpeg.Canvas.Font.Bold	=	True
	Jpeg.Canvas.Font.BkMode =	&HFFFFFF
	Jpeg.Canvas.Font.Quality=	10
	Jpeg.Canvas.Print JpegLeft, JpegTop, PrintWords
	Jpeg.Width = JWidth
	'Jpeg.crop 0,0,70,52
	JpegName = HomeElitePath & GetRndNum(4) & ".jpg"
	Jpeg.Save Server.MapPath("../" & JpegName)
	Set Jpeg = Nothing
	Response.Write JpegName
ElseIf Request("Action") = "delete" Then
	Dim Image
	Image = Request("Image")
	Call DelImage(Image)
End If


'**************************************************
'函数名：GetRndNum
'作  用：产生制定位数的随机数
'参  数：iLength ---- 随即数的位数
'返回值：随机数
'**************************************************
Function GetRndNum(iLength)
    Dim i, str1
    For i = 1 To (iLength \ 5 + 1)
        Randomize
        str1 = str1 & CStr(CLng(Rnd * 90000) + 10000)
    Next
    GetRndNum = Year(Now) & Right("0"&Month(Now),2) & Right("0"&Day(Now),2) & Right("0"&Hour(Now),2) & Right("0"&Minute(Now),2) & Right("0"&Second(Now),2) & Left(str1, iLength)
End Function

'**************************************************
'函数名：GetStrLen
'作  用：求字符串长度。汉字算两个字符，英文算一个字符。
'参  数：str  ----要求长度的字符串
'返回值：字符串长度
'**************************************************
Function GetStrLen(str)
    On Error Resume Next
    Dim WINNT_CHINESE
    WINNT_CHINESE = (Len("中国") = 2)
    If WINNT_CHINESE Then
        Dim l, t, c
        Dim i
        l = Len(str)
        t = l
        For i = 1 To l
            c = Asc(Mid(str, i, 1))
            If c < 0 Then c = c + 65536
            If c > 255 Then
                t = t + 1
            End If
        Next
        GetStrLen = t
    Else
        GetStrLen = Len(str)
    End If
    If Err.Number <> 0 Then Err.Clear
End Function

'**************************************************
'函数名：DelImage
'作  用：删除图片
'参  数：url  ---- 图片地址
'返回值：
'**************************************************
Function DelImage(url) 
	On Error Resume Next 
	Dim FSO 
	Set FSO = Server.CreateObject("Scripting.FileSystemObject") 
	FSO.DeleteFile Server.MapPath("../" & url),true
	Set FSO = Nothing
	Response.Write Server.MapPath("../" & url) & " Deleted."
End Function 

%>