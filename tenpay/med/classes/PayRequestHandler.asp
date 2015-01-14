<!--#include file="../util/md5.asp"-->
<!--#include file="../util/tenpay_util.asp"-->
<%
'
'��ʱ����֧��������
'============================================================================
'api˵����
'init(),��ʼ��������Ĭ�ϸ�һЩ������ֵ����cmdno,date�ȡ�
'getGateURL()/setGateURL(),��ȡ/������ڵ�ַ,����������ֵ
'getKey()/setKey(),��ȡ/������Կ
'getParameter()/setParameter(),��ȡ/���ò���ֵ
'getAllParameters(),��ȡ���в���
'getRequestURL(),��ȡ������������URL
'doSend(),�ض��򵽲Ƹ�֧ͨ��
'getDebugInfo(),��ȡdebug��Ϣ
'
'============================================================================
'

Class PayRequestHandler
	
	'����url��ַ
	Private gateUrl
	
	'��Կ
	Private key
	
	'����Ĳ���
	Private parameters
	
	'debug��Ϣ
	Private debugInfo
	
	'��ʼ���캯��
	Private Sub class_initialize()
		gateUrl = "https://gw.tenpay.com/gateway/pay.htm"
		key = ""
		Set parameters = Server.CreateObject("Scripting.Dictionary")
		debugInfo = ""
	End Sub
	
	'��ʼ������
	Public Function init()
		parameters.RemoveAll
	End Function
	
	'��ȡ��ڵ�ַ,����������ֵ
	Public Function getGateURL()
		getGateURL = gateUrl
	End Function
	
	'������ڵ�ַ,����������ֵ
	Public Function setGateURL(gateUrl_)
		gateUrl = gateUrl_
	End Function
	
	'��ȡ��Կ
	Public Function getKey()
		getKey = key
	End Function
	
	'������Կ
	Public Function setKey(key_)
		key = key_
	End Function
	
	'��ȡ����ֵ
	Public Function getParameter(parameter)
		getParameter = parameters.Item(parameter)
	End Function
	
	'���ò���ֵ
	Public Sub setParameter(parameter, parameterValue)
		If parameters.Exists(parameter) = True Then
			parameters.Remove(parameter)
		End If
		parameters.Add parameter, parameterValue	
	End Sub

	'��ȡ��������Ĳ���,����Scripting.Dictionary
	Public Function getAllParameters()
		getAllParameters = parameters
	End Function
	
	'��ȡ������������URL
	Public Function getRequestURL()

		Call createSign()
		
		Dim reqPars
		Dim k
		For Each k In parameters
			reqPars = reqPars & k & "=" & Server.URLEncode(parameters(k)) & "&" 
		Next
		
		'ȥ�����һ��&
		reqPars = Left(reqPars, Len(reqPars)-1)

		getRequestURL = getGateURL & "?" & reqPars

	End Function
	
	'�ض��򵽲Ƹ�֧ͨ��
	Public Function doSend()
		Response.Redirect(getRequestURL())
		Response.End
	End Function	
	
	'��ȡdebug��Ϣ
	Public Function getDebugInfo()
		getDebugInfo = debugInfo
	End Function
	
	'����ǩ��
	Private Sub createSign()
		Dim keys,k,v
		keys	= parameters.Keys()
		'����ĸ˳������
		for i=0 to UBound(keys)-1
			for j=i+1 to UBound(keys)
				if StrComp(keys(i), keys(j)) > 0 then 
					tmp=keys(i)
					keys(i)=keys(j)
					keys(j)=tmp
				end if
			next
		next
		md5str = ""
		'���ǩ���ַ���
		For Each k in keys
			v = getParameter(k)
			if v <> "" and k <> "sign" and k <> "key" then
				md5str = md5str & k & "=" & v & "&"
			end if
		Next

		md5str = md5str & "key=" & key

		Dim sign
		sign= LCase(ASP_MD5(md5str))

		setParameter "sign", sign

		'debuginfo
		debugInfo = md5str & " => sign:" & sign
		
	End Sub

End Class

%>