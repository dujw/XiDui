<%
	if kaiguan = "0" then 
		response.Write "<br><center>"&guanbi&"</font>"
		response.End()
	end if 
%>
<table width="950" height="70" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="210" rowspan="2"><a href="http://www.dk501.com"><img src="image/logo/logo.jpg" width="180" height="45" border="0" /></a></td>
    <td height="38" colspan="2" align="right">�ҵ��ʻ� | ������֪ | ��Ա���� | ����˵�� | ����˿� | ��λ��� | ��ϵ����</td>
  </tr>
  <tr>
    <td width="583"><% if request.cookies("buyok")("userid")="" then%>
      ���ã���ӭ���ٲ��У�<a href="login.asp">[���½]</a> <a href="/login/reg.asp">[���ע��]</a>
      <%else
	response.Write "��ã�" & request.cookies("buyok")("userid") & "<a href='/login/logout.asp'>[�˳�]<a> <a href='/user/user_center.asp'>�ҵĶ���</a> "
	end if%>    </td>
    <td width="157" align="right" class="red"><img src="images/index/tel.jpg" width="232" height="20"></td>
  </tr>
</table>
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30" colspan="5">
	<table border="0" cellpadding="0" cellspacing="0">
      <tr>
	  <%
	LarCode=trim(request("LarCode"))'��Ʒ�����
	MidCode=trim(request("MidCode"))'��ƷС����
	if LarCode = "" then
		response.Write "<td  width=85 height=30 align=center background=images/index/top1.jpg><a href=index.asp class=titile>��վ��ҳ</a></td><td width=2></td>"
	else
		response.Write "<td  width=85 height=30 align=center background=images/index/top1.jpg><a class=text href=index.asp>��վ��ҳ</a></td><td width=2></td>"
	end if
	set rs_product=server.createobject("adodb.recordset") 
	sql_product="select distinct LarCode,LarSeq from dk501_productclass order by LarSeq" 
	rs_product.Open sql_product,conn,1,1 		
	if rs_product.bof and rs_product.eof then
		response.write "�������ڸ����У������ĵȺ�"
	else
		Do While Not rs_product.eof	
		if LarCode =  Trim(rs_product("LarCode")) then 
			response.write "<td  width=85 height=30 align=center background=images/index/top1.jpg><a class=titile href=product.asp?LarCode="&Server.URLEncode(Trim(rs_product("LarCode")))&">"&Trim(rs_product("LarCode"))&"</a></td><td width=2></td>"
		else
			response.write "<td  width=85 height=30 align=center background=images/index/top1.jpg><a class=text href=product.asp?LarCode="&Server.URLEncode(Trim(rs_product("LarCode")))&">"&Trim(rs_product("LarCode"))&"</a></td><td width=2></td>"
		end if 
		rs_product.movenext
		  if rs_product.eof then exit do
		loop		
	end if
	rs_product.close
	set rs_product=nothing
	%>
      </tr>
    </table></td>
  </tr>
  <% if trim(request("LarCode")) <> "" then%>
  <tr>
    <td height="20" colspan="5" bgcolor="#565662">
	
	<%
	set rs_MidCode=server.createobject("adodb.recordset") 
	sql_MidCode="select TOP 12 LarCode,MidCode from dk501_productclass order by MidSeq" 
	rs_MidCode.Open sql_MidCode,conn,1,1 	
	if rs_MidCode.bof and rs_MidCode.eof then
		response.write "���޷��࣡"
	else
		Do While Not rs_MidCode.eof	
			response.write "<a  href=product.asp?LarCode="&Server.URLEncode(Trim(rs_MidCode("LarCode")))&"&MidCode="&Server.URLEncode(Trim(rs_MidCode("MidCode")))&">"&Trim(rs_MidCode("MidCode"))&"</a>&nbsp;&nbsp;"
		rs_MidCode.movenext
		  if rs_MidCode.eof then exit do
		loop		
	end if
	rs_MidCode.close
	set rs_MidCode=nothing
	%>	</td>
  </tr>
  <% end if %>
  <tr>
    <td width="9" height="30" bgcolor="#F5F5F5" class="b-top">&nbsp;</td>
    <td width="124" bgcolor="#F5F5F5" class="b-top"><%
	set rs_1=server.createobject("adodb.recordset") 
	sql="select LarCode from dk501_productclass where MidSeq = 1 order by LarSeq asc" 
	rs_1.Open sql,conn,1,1 
		response.write "<select name='fenlei'>"
		response.write	"<option value='' selected>�������</option>"
	if rs_1.bof and rs_1.eof then
		response.write "<option selected value=''>�������</option>"
	else
		Do While Not rs_1.eof	
		response.write "<option value='"&rs_1("LarCode")&"'> == "&rs_1("LarCode")&" == </option>"
		rs_1.movenext
		  if rs_1.eof then exit do
		loop
		response.write "<option value='δ����'> == ���� == </option>"
	end if
	rs_1.close
	set rs_1=nothing
	%></td>
    <td width="304" bgcolor="#F5F5F5" class="b-top"><input name="name_key" type="text" id="name_key" size="40"></td>
    <td width="106" bgcolor="#F5F5F5" class="b-top"><input name="image" type="image" src="images/index/serch.gif"/></td>
    <td width="407" bgcolor="#F5F5F5" class="b-top">���Źؼ���:���� ���п� ���� POLO/T�� ȹ�� ţ�п� Χ��  </td>
  </tr>
  <tr>
    <td height="10" colspan="5"></td>
  </tr>
</table>
