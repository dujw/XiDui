<div id='header_login'>
  <ul>
    <li>
      <% if request.cookies("buyok")("userid")="" then%>
      您好！<a href="/login/login.asp">[请登陆]</a> <a href="/login/reg.asp">[免费注册]</a>
      <%else
	response.Write "你好，" & request.cookies("buyok")("userid") & "<a href='/login/logout.asp'>[退出]<a> <a href='/user/user_center.asp'>我的订单</a> "
	end if%>
      | <a href="/user/index.asp">用户中心</a> | <a href="/Contact/Contact.asp?id=13">我要加盟</a> | <a href="/Contact/Contact.asp?id=14">会员服务</a> | <a href="/Contact/Contact.asp?id=15">在线客服</a> | <a href="/Contact/Contact.asp?id=12">帮助中心</a></li>
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
		response.write	"<option value='' selected>所有类别</option>"
	if rs_1.bof and rs_1.eof then
		response.write "<option selected value=''>暂无类别</option>"
	else
		Do While Not rs_1.eof	
		response.write "<option value='"&rs_1("LarCode")&"'> == "&rs_1("LarCode")&" == </option>"
		rs_1.movenext
		  if rs_1.eof then exit do
		loop
		response.write "<option value='未归类'> == 其它 == </option>"
	end if
	rs_1.close
	set rs_1=nothing
	%>
    <input type="submit" name="button32" id="button32" value="搜索" class="search-input"/>
  </div>
  </form>
  <div id='header_login_logo_right'>
    <P>高级搜索</P>
    <P>使用帮助</P>
  </div>
</div>
