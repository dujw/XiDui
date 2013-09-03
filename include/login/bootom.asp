<div id='div_Public'>
<div id='foot_link'><span>合作伙伴：</span>
  <%
Set rs_links = conn.Execute("select url,site from dk501_links where online='1' order by txt desc,num asc") 
if not (rs_links.eof and rs_links.bof) then
	Do while not rs_links.eof
	response.write "<a href='"&rs_links("url")&"' target='_blank'>"&lleft(rs_links("site"),12)&"</a>&nbsp;"
	rs_links.movenext
	loop
end if     
rs_links.close
set rs_links=nothing
%>
</div>
<div id='foot_login'>
<a href="/Contact/Contact.asp?id=1">关于我们</a> 
<a href="/Contact/Contact.asp?id=2">付款说明</a> 
<a href="/Contact/Contact.asp?id=3">普通快递</a> 
<a href="/Contact/Contact.asp?id=4">配送细节</a> 
<a href="/Contact/Contact.asp?id=6">联系我们</a> 
<a href="/Contact/Contact.asp?id=7">如何换货</a> 
<a href="/Contact/Contact.asp?id=8">代理须知</a> 
<a href="/Contact/Contact.asp?id=9">如何退款</a> 
<a href="/Contact/Contact.asp?id=10">批发咨询</a> 
<a href="/Contact/Contact.asp?id=12">帮助中心</a> 
<br>
  Copyright 2008-2010 Tencent.All Rights Reserved , 版权所有 WWW.DK501.COM 裁男<br>
  中国信息产业部备案号：渝ICP备09003182号
</div>
</div>
