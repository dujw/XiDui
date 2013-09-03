<%
	if kaiguan = "0" then 
		response.Write "<br><center>"&guanbi&"</font>"
		response.End()
	end if 
%>
<%
		SSQL="select top 1 id,Ntitle,NEditTime from DU_Notice where NOverTime>getdate() order by id desc"
'response.write SSQL
set djw_rs=Server.Createobject("ADODB.RecordSet")
djw_rs.open SSQL,conn,1,3
if djw_rs.eof and djw_rs.bof  then
	response.write "<!--by djw 2012.2.25-->"
else
%>
<script type="text/javascript" src="js_lib/djw/lhgcore.lhgdialog.min.js?self=true&skin=mac"></script>
<script type="text/javascript">
function dk_notice(){
	$.dialog({width:700,height:280,opacity:0.5,background:'#ccc',lock:true,title: '裁难 <%=djw_rs(2)%> 发布',content: 'url:Notice.asp?id=<%=djw_rs(0)%>',button: [{ name: '关闭'}],padding:5});
}
</script>
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0" style=" border:#CCCCCC solid 1px;">
  <tr >
    <td height="30" align="left"><strong><center><a href="#" onclick="dk_notice()">裁男最新通知:<font color="#FF0000"><%=djw_rs(1)%></font>&nbsp;&nbsp;[<%=djw_rs(2)%>]</a></center></strong>
</td>
  </tr>
</table>
<!--<table width="950" border="0" align="center" cellpadding="0" cellspacing="0" style=" border:#CCCCCC solid 0px;">
  <tr >
    <td height="30" align="left"><img src="" border="0" />
	</td>
  </tr>
</table>-->
<%
djw_rs.close()
end if%>
<table width="950" height="70" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="180" rowspan="2"><a href="http://www.dk501.com"><img src="../image/logo/logo.png" width="180" height="45" border="0" /></a></td>
    <td height="38" colspan="2" align="right"><% if request.cookies("buyok")("userid")="" then%>
      您好，欢迎光临裁男！<a href="../login.asp">[请登陆]</a> <a href="../reg.asp">[免费注册]</a>
      <%else
	response.Write "你好，" & request.cookies("buyok")("userid") & "<a href='../login/logout.asp'>[退出]<a> <a href='../user_center.asp'>我的订单</a>  <a href='../check.asp'>购物车</a>"
	end if%>     <a href="../user_index.asp">我的帐户</a> | <a href="../Contact.asp?id=8">代理须知</a> | <a href="../Contact.asp?id=14">会员服务</a> | <a href="../Contact.asp?id=2">付款说明</a> | <a href="../Contact.asp?id=9">如何退款</a> | <a href="../Contact.asp?id=7">如何换货</a> | <a href="../Contact.asp?id=6">联系我们</a></td>
  </tr>
  <tr>
    <td width="537" align="right">
    <a href="../user_sjb.asp"><font color="#FF0000" size="+1"><strong>数据包下载地址</strong></font></a>
    <a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=1504490396&site=裁难小刘&menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=2:1504490396:45" alt="点击这里给我发消息" title="点击这里给我发消息">裁男小刘</a>
	<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=2262462307&site=裁难小周&menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=2:2262462307:45" alt="点击这里给我发消息" title="点击这里给我发消息">裁男小周</a>
	<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=2404470991&site=裁难小杨&menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=2:2404470991:45" alt="点击这里给我发消息" title="点击这里给我发消息">裁男小杨</a>
    <a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=2486488687&site=裁难小许&menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=2:2486488687:45" alt="点击这里给我发消息" title="点击这里给我发消息">裁男小许</a>
	</td>
    <td width="233" align="right" class="red"><img src="../images/index/tel.png" width="232" height="20"></td>
  </tr>
</table>
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30" colspan="5">
	<table border="0" cellpadding="0" cellspacing="0">
      <tr>
	  <%
	LarCode=trim(request("LarCode"))'产品大分类
	MidCode=trim(request("MidCode"))'产品小分类
	if LarCode = "" then
		response.Write "<td  width=85 height=30 align=center background=images/index/top1.jpg><a href=../index.asp class=titile_1>网站首页</a></td><td width=2></td>"
	else
		response.Write "<td  width=85 height=30 align=center background=images/index/top1.jpg><a class=text_1 href=../index.asp>网站首页</a></td><td width=2></td>"
	end if
	set rs_product=server.createobject("adodb.recordset") 
	sql_product="select distinct LarCode,LarSeq from dk501_productclass order by LarSeq" 
	rs_product.Open sql_product,conn,1,1 		
	if rs_product.bof and rs_product.eof then
		response.write "数据正在更新中，请耐心等候！"
	else
		Do While Not rs_product.eof	
		if LarCode =  Trim(rs_product("LarCode")) then 
			response.write "<td  width=85 height=30 align=center background=images/index/top1.jpg><a class=titile_1 href=../product.asp?LarCode="&Server.URLEncode(Trim(rs_product("LarCode")))&">"&Trim(rs_product("LarCode"))&"</a></td><td width=2></td>"
		else
			response.write "<td  width=85 height=30 align=center background=images/index/top1.jpg><a class=text_1 href=../product.asp?LarCode="&Server.URLEncode(Trim(rs_product("LarCode")))&">"&Trim(rs_product("LarCode"))&"</a></td><td width=2></td>"
		end if 
		rs_product.movenext
		  if rs_product.eof then exit do
		loop		
	end if
	rs_product.close
	set rs_product=nothing
	response.Write "<td  width=85 height=30 align=center background=images/index/top1.jpg><a class=text_1 href='xuanchuan.asp'><span style='color:#F00; font-weight:bold;'>PSD源文件</span></a></td><td width=2></td>"
	%>
    
      </tr>
    </table></td>
  </tr>
  <% if trim(request("LarCode")) <> "" then%>
  <tr>
    <td height="30" colspan="5" bgcolor="#CCCCCC">&nbsp;→&nbsp;
	<%
	SubLar=trim(request("LarCode"))
	set rs_MidCode=server.createobject("adodb.recordset") 
	sql_MidCode="select LarCode,MidCode from dk501_productclass where LarCode = '"& SubLar &"'order by MidSeq" 
	rs_MidCode.Open sql_MidCode,conn,1,1 	
	if rs_MidCode.bof and rs_MidCode.eof then
		response.write "暂无分类！"
	else
		Do While Not rs_MidCode.eof	
			response.write "<a href=../product.asp?LarCode="&Server.URLEncode(Trim(rs_MidCode("LarCode")))&"&MidCode="&Server.URLEncode(Trim(rs_MidCode("MidCode")))&">"&Trim(rs_MidCode("MidCode"))&"</a>&nbsp;&nbsp;"
		rs_MidCode.movenext
		  if rs_MidCode.eof then exit do
		loop		
	end if
	rs_MidCode.close
	set rs_MidCode=nothing
	%></td>
  </tr>
<% end if %>
   <form action="product.asp" method="get" name="search" id="search">
  <tr>
    <td width="9" height="30" bgcolor="#F5F5F5" class="b-top">&nbsp;</td>
    <td width="124" bgcolor="#F5F5F5" class="b-top"><%
	set rs_1=server.createobject("adodb.recordset") 
	sql="select LarCode from dk501_productclass where MidSeq = 1 order by LarSeq asc" 
	rs_1.Open sql,conn,1,1 
		response.write "<select name='fenlei' >"
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
	%></td>
    <td width="304" bgcolor="#F5F5F5" class="b-top"><input name="name_key" type="text" id="name_key" size="40" value="<%=trim(request("name_key"))%>"></td>
    <td width="106" bgcolor="#F5F5F5" class="b-top"><input name="image" type="image" src="../images/index/serch.gif"/></td>
    <td width="407" bgcolor="#F5F5F5" class="b-top"></td>
  </tr>
  </form>
  <tr>
    <td height="10" colspan="5"></td>
  </tr>
</table>
