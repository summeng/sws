<%

'��ȡʱ���ַ���, ��ʽYYYYMMDDhhmiss
Public Function getStrNow()
	strNow = Now()
	strNow = Year(strNow) & Right(("00" & Month(strNow)),2) & Right(("00" & Day(strNow)),2) & Right(("00" & Hour(strNow)),2) & Right(("00" &  Minute(strNow)),2) & Right(("00" & Second(strNow)),2)
	getStrNow = strNow
End Function


' ��ȡ���������ڣ���ʽYYYYMMDD
Public Function getServerDate()
	getServerDate = Left(getStrNow,8)
End Function

'��ȡʱ��,��ʽhhmiss ��:192030
Public Function getTime()
	getTime = Right(getStrNow,6)
End Function

'��ȡ�����,���� [min,max]��Χ����
Public Function getRandNumber(max, min)
	Randomize 
	getRandNumber = CInt((max-min+1)*Rnd()+min) 
End Function

'��ȡ������ֵ��ַ���,����[min,max]��Χ�������ַ���
Public Function getStrRandNumber(max, min)
	randNumber = getRandNumber(max, min)
	getStrRandNumber = CStr(randNumber)
End Function

'д��־��������ԣ�����վ����Ҳ���Ըĳɴ������ݿ⣩
'sWord Ҫд����־����ı�����
Public Function log_result(sWord)
	set fs= createobject("scripting.filesystemobject")
	set ts=fs.createtextfile(server.MapPath("/tenpay/b2c/log/"&getServerDate()&getTime()&"."&getStrRandNumber(1000,9999)&".txt"),true)
	ts.writeline(sWord)
	ts.close
	set ts=Nothing
	set fs=Nothing
End Function


%>