<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Private Function Html2Js(str)
    If str <> "" Then
        str = Replace(str, Chr(34), "\""")
        str = Replace(str, Chr(39), "\'")
        str = Replace(str, Chr(13), "\n")
        str = Replace(str, Chr(10), "\r")
    End If
    Html2Js = str
End Function%>