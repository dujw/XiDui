<div id='header_login'>
  <ul>
    <li>
      <% if request.cookies("buyok")("userid")="" then%>
      ���ã�<a href="/login/login.asp">[���½]</a> <a href="/login/reg.asp">[���ע��]</a>
      <%else
	response.Write "��ã�" & request.cookies("buyok")("userid") & "<a href='/login/logout.asp'>[�˳�]<a> <a href='/user/user_center.asp'>�ҵĶ���</a> "
	end if%>
      | <a href="/user/index.asp">�û�����</a> | <a href="/Contact/Contact.asp?id=13">��Ҫ����</a> | <a href="/Contact/Contact.asp?id=14">��Ա����</a> | <a href="/Contact/Contact.asp?id=15">���߿ͷ�</a> | <a href="/Contact/Contact.asp?id=12">��������</a></li>
  </ul>
</div>
<div id='header_login_logo'>
  <div id='header_login_logo_left'><img src="../image/logo/logo.jpg" width="180" height="45" /> </div>
  <form action="/product/product.asp" method="get" name="search" id="search">
  <div id='header_login_logo_mid'>
    <input name="name_key" type="text" class="input_search" id="name_key" size="35">
    <%
	set rs_1=server.createobject("adodb.recordset") 
	sql="select LarCode from dk501_productclass where MidSeq = 1 order by LarSeq asc" 
	rs_1.Open sql,conn,1,1 
		response.write "<select name='fenlei' class=select_search>"
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
	%>
    <input type="submit" name="button32" id="button32" value="����" class="search-input"/>
  </div>
  </form>
  <div id='header_login_logo_right'>
    <P>�߼�����</P>
    <P>ʹ�ð���</P>
  </div>
</div>
