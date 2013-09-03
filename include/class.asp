<table width="200" border="0" cellpadding="0" cellspacing="1" bgcolor="#B4B4B4">
      <tr>
        <td height="25" bgcolor="#565662" style="padding:6px;"><font color="#FFFFFF"><strong>产品分类</strong></font></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF" style="padding:6px;">
		<%
openlar=request("larcode")
Set Fso = CreateObject("Scripting.FileSystemObject")
set buyok_class=Fso.openTextFile(Server.MapPath("manage/include/class.asp"))
do while not buyok_class.atendofstream
classtext=buyok_class.readline
if classtext="暂无商品分类" then
	response.write "暂无商品分类"
exit do
end if
if instr(classtext,"lar:")>0 then
larcode=replace(classtext,"lar:","")
if tree_num=1 then
	thislarnum=cqrsprod("thislarnum")
	thislarnum="["&thislarnum&"]"
	set cqrsprod=nothing
end if
lar=lar+1
if lar>1 and lar<>old then response.write "<br>"
	midd=1
	response.write "<a class='titile' href='product.asp?LarCode="&server.URLEncode(LarCode)&"'>■ "&larcode&" ■</a>"
else
	midcode=classtext
	if midd=1 then
	response.write "<br>"
end if
response.write "&nbsp;→<a href='product.asp?LarCode="&server.URLEncode(LarCode)&"&MidCode="&server.URLEncode(MidCode)&"'>&nbsp;"&lleft(MidCode,20)&"</a><br>"
midd=midd+1
old=lar
end if
loop
buyok_class.close
set fso=nothing
%>
		</td>
      </tr>
    </table>