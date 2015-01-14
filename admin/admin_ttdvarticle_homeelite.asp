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
'��������GetRndNum
'��  �ã������ƶ�λ���������
'��  ����iLength ---- �漴����λ��
'����ֵ�������
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
'��������GetStrLen
'��  �ã����ַ������ȡ������������ַ���Ӣ����һ���ַ���
'��  ����str  ----Ҫ�󳤶ȵ��ַ���
'����ֵ���ַ�������
'**************************************************
Function GetStrLen(str)
    On Error Resume Next
    Dim WINNT_CHINESE
    WINNT_CHINESE = (Len("�й�") = 2)
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
'��������DelImage
'��  �ã�ɾ��ͼƬ
'��  ����url  ---- ͼƬ��ַ
'����ֵ��
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