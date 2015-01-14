<%

'获取时间字符串, 格式YYYYMMDDhhmiss
Public Function getStrNow()
	strNow = Now()
	strNow = Year(strNow) & Right(("00" & Month(strNow)),2) & Right(("00" & Day(strNow)),2) & Right(("00" & Hour(strNow)),2) & Right(("00" &  Minute(strNow)),2) & Right(("00" & Second(strNow)),2)
	getStrNow = strNow
End Function


' 获取服务器日期，格式YYYYMMDD
Public Function getServerDate()
	getServerDate = Left(getStrNow,8)
End Function

'获取时间,格式hhmiss 如:192030
Public Function getTime()
	getTime = Right(getStrNow,6)
End Function

'获取随机数,返回 [min,max]范围的数
Public Function getRandNumber(max, min)
	Randomize 
	getRandNumber = CInt((max-min+1)*Rnd()+min) 
End Function

'获取随机数字的字符串,返回[min,max]范围的数字字符串
Public Function getStrRandNumber(max, min)
	randNumber = getRandNumber(max, min)
	getStrRandNumber = CStr(randNumber)
End Function

'写日志，方便测试（看网站需求，也可以改成存入数据库）
'sWord 要写入日志里的文本内容
Public Function log_result(sWord)
	set fs= createobject("scripting.filesystemobject")
	set ts=fs.createtextfile(server.MapPath("/tenpay/b2c/log/"&getServerDate()&getTime()&"."&getStrRandNumber(1000,9999)&".txt"),true)
	ts.writeline(sWord)
	ts.close
	set ts=Nothing
	set fs=Nothing
End Function


%>