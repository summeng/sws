<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=top.asp-->

<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>

<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
</head>


<table width="980" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="980" valign="top"><table width="980" border="0" cellpadding="0" cellspacing="0"  bgcolor="#ffffff">
      <tr>
        <td height="414" valign="top">   <div>
            <br />
            <table width="100%" height="354" border="0" align="center" cellpadding="0" cellspacing="1" class="PayPut1">
              <tr valign="middle">
                <td height="25" nowrap="nowrap" bgcolor="#EDDAFF"><div class="style36"> <span class="aa  TopicShow5"> &nbsp;&gt;&gt; ������Ϣ��ʾ������ѡ�񷢲����� </span></div></td>
              </tr>
			  <tr><td width=980 height=120 align="center"><img src=img/reg.jpg border=0></td></tr>
              <tr valign="middle">
                <td height="40" nowrap="nowrap"><table width="543" height="283"  border="0" align="center" cellpadding="0" cellspacing="16">
                  <tr>
                    <td width="198" height="45" background="img/cardTop.gif"><div align="center" class="style34 STYLE31 STYLE31"><a href="<%=addinfo%>?<%=CT_ID%>&ttop=2"><span style="color:#FF6600 ">�οͷ���</span></a></div></td>
                    <td width="83"  >&nbsp;</td>
                    <td width="198" background="img/cardTop.gif"><div align="center" class="style34" ><a href="<%=reg%>?<%=CT_ID%>&ttop=10"><span style="color:#FF6600 ">ע�ᷢ��</span></a></div></td>
                    <td width="83"  >&nbsp;</td>
                    <td width="198" background="img/cardTop.gif"><div align="center" class="style34" ><a href="help.asp?id=1&<%=CT_ID%>&ttop=10"><span style="color:#FF6600 ">VIP����</span></a></div></td>
                  </tr>
                  <tr>
                    <td width="198" height="126" valign="top" bgcolor="#F7FFFF" class="tb1"><span style="line-height:200%;margin-left:5px"><span>
                      <table width="180"  border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr>
                          <td height="10"></td>
                        </tr>
                        <tr>
                          <td>
                            <li>�οͷ�����Ϣ<%If onoff =1 And xinxish=1 then%>�����<%else%>�����<%End if%></li>
                            <li>�޻��ֻ���</li>
                            <li>����Ա��ʱ������ɾ����Ϣ</li>
                            <li>���ö����ܣ���վ�����Ƽ�</li>
                            <li>�����ޱ�ɫ�Ĺ���</li>
							<li>���ϴ�ͼƬ���ܣ���ͼ��Ч</li>
                            <li>������Ϣʱ���ݿ��HTM�༭��</li>
                            <li>�ѷ�����Ϣ�����޸ĺ�ɾ��</li>
							<li>������Ϣ���ܰ��û�������</li>
							<li>�������뿪ͨ�������ҵ��Ƭ</li>
                            <li>ÿ�η�����Ϣ��Ҫ��������Ϣ</li></td>
                        </tr>
                      </table>
                    </span></span> </td>
                    <td width="83" valign="top"  ></td>
                    <td width="198" valign="top" bgcolor="#F7FFFF" class="tb1"><span style="line-height:200%;margin-left:5px">
                      <table width="180"  border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr>
                          <td height="10"></td>
                        </tr>
                        <tr>
                          <td>
                            <li>��ͨ��Ա������Ϣ<%If onoff =1 And vipno=1 then%>�����<%else%>�����<%End if%></li>
                            <li>�л��ֻ���</li>
                            <li>��Ϣ�����ڲ�Υ�治�ᱻɾ��</li>
                            <li>���ö����ܣ���վ�����Ƽ�</li>
                            <li>����ɱ�ɫ��ͻ����ʾ</li>
							<li>���ϴ�ͼƬһ�ţ�����ͼ����</li>
                            <li>������Ϣʱ���ݿ�ΪHTM�༭��</li>
                            <li>�ѷ�����Ϣ����ʱ�޸�ɾ��</li>
							<li>��Ϣ�ɰ������ģ�������ʾ<%=huiyuanxx%>��</li>
							<li>�ɿ�������Ƭ,��Ϣ����ʾ<%=huiyuansp%>��</li>
                            <li>ÿ�η�����Ϣ������������Ϣ</li></td>
                        </tr>
                      </table>
                    </span></td>
                    <td width="83" valign="top"  ></td>
                    <td width="198" valign="top" bgcolor="#F7FFFF" class="tb1"><span style="line-height:200%;margin-left:5px">
                      <table width="180"  border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr>
                          <td height="10"></td>
                        </tr>
                        <tr>
                          <td>
                            <li>VIP��Ա������Ϣ<%If onoff=1 then%>�����<%else%>�����<%End if%></li>
                            <li>�л��ֻ���</li>
                            <li>��Ϣ�����ڲ�Υ�治�ᱻɾ��</li>
                            <li>���ö����ܣ��ɻ�ñ�վ�Ƽ�</li>
                            <li>����ɱ�ɫ��ͻ����ʾ</li>
							<li>���ϴ�ͼƬ��������ͼȨ��</li>
                            <li>������Ϣʱ���ݿ�ΪHTM�༭��</li>
                            <li>�ѷ�����Ϣ����ʱ�޸�ɾ��</li>
							<li>��Ϣ�ɰ������ģ�������ʾ</li>
							<li>�ɿ�������Ƭ,��Ϣ������ʾ</li>
                            <li>ÿ�η�����Ϣ������������Ϣ</li></td>
                        </tr>
                      </table>
                    </span></td>
                  </tr>
                  <tr valign="middle">
                    <td height="30" ><div align="center"><a href="<%=addinfo%>?<%=CT_ID%>&ttop=2"><img src="img/PostSelect1.gif"   border="0" /></a> </div></td>
                    <td><div align="center"> </div></td>
                    <td><div align="center"><a href="<%=reg%>?<%=CT_ID%>&ttop=10"><img src="img/PostSelect2.gif" border="0"   /></a> </div></td>
                  <td><div align="center"> </div></td>
                    <td><div align="center"><a href="help.asp?id=1&<%=CT_ID%>&ttop=10"><img src="img/PostSelect3.gif" border="0"   /></a> </div></td>
                  </tr>
                </table>                  </td>

              </tr>
            </table> 
            <br />
<center>
<!--JS���÷�ʽ�����뿪ʼ-->
  <script src="<%=adjs_path("ads/js","syzb730",c1&"_"&c2&"_"&c3)%>"></script>
  <!--���������-->
      
            </div></td>
        </tr>
    </table>	</td>
  </tr>
</table>
 
   <table width="980" cellpadding="0" cellspacing="0" align="center" >
     <tr>
       <td><!--#include file=footer.asp--></td>
     </tr>
</table>
</body>
</html>
