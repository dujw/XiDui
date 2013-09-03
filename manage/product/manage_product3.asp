<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<%
'修改商品资料
if request("modify")="edit" then
	if request("midcode")="" then
		response.write "<script language='javascript'>"
		response.write "alert('您没有选择有效的商品分类，单击“确定”返回上一页！');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
	response.end
	end if
if request("ProdId")=""  or request("Prodname")="" or request("PriceList")="" or request("PriceOrigin")="" then
	response.write "<script language='javascript'>"
	response.write "alert('商品资料填写不完整，单击“确定”返回上一页！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end if

if len(request("ProdDisc"))>200 then
response.write "<script language='javascript'>"
response.write "alert('您填写的商品简介内容太长了，单击“确定”返回上一页！');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if

id=Trim(request("ProdId"))
if id="" then
	response.write "<script language='javascript'>"
	response.write "alert('商品不存在，请检查！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end if
set rs=server.createobject("adodb.recordset")
sql = "select * from dk501_product where ProdId='"&id&"'"
rs.open sql,conn,1,3
if not (rs.eof and rs.bof) then
	rs("ProdId")=Trim(rs("ProdId"))
	rs("ProdName")=Trim(request.form("ProdName"))'商品名称
	rs("prod1")=Trim(request.form("prod1"))'自定义一	
	rs("zhongliang")=Trim(request.form("zhongliang"))
	rs("PriceOrigin")=Trim(request.form("PriceOrigin"))'原价
	rs("PriceList")=Trim(request.form("PriceList"))'现价
	code=split(request.form("Midcode"),"^")
	rs("LarCode")=code(0)'大类
	rs("MidCode")=code(1)'中类
	rs("PriceList1")=request.form("PriceList1")'现价
	rs("PriceList2")=request.form("PriceList2")'现价
	rs("PriceList3")=request.form("PriceList3")'现价
	rs("PriceList4")=request.form("PriceList4")'现价
	rs("PriceList5")=request.form("PriceList5")'现价
	rs("PriceList6")=request.form("PriceList6")'现价
	rs("jifenzhi")=request.form("jifenzhi")'积分
	rs("Product_Price")=Int(request.form("Product_Price"))'产品价格
	rs("Product_number")=Trim(request.form("Product_number"))'产品货号
	rs("Product_Factory")=request.form("Product_Factory")'产品厂家
	rs("ProdDisc")=Trim(request.form("ProdDisc"))'简介
	rs("ImgPrev")=Trim(request.form("ImgPrev"))'商品图片
	rs("ImgPrev2")=Trim(request.form("ImgPrev2"))'商品图片
	rs("MemoSpec")=Trim(request.form("MemoSpec"))'商品介绍
	
	MemoSpec_AD=request.form("MemoSpec_AD")	'细节广告图文件路径
	if MemoSpec_AD="" then
		MemoSpec_AD=null
	end if
	rs("MemoSpec_AD")=MemoSpec_AD
	
	MemoSpec_PSD=request.form("MemoSpec_PSD")	'PSD源文件路径
	if MemoSpec_PSD="" then
		MemoSpec_PSD=null
	end if
	rs("MemoSpec_PSD")=MemoSpec_PSD
	
	rs("Remark")=request.form("Remark")'是否推荐
	rs("Quantity")=request.form("Quantity")
	if request.form("up")="1" then rs("TJDate")=now()'是否提升
rs.update
rs.close
set rs=nothing
end if
response.write "<script language='javascript'>"
response.write "alert('您已成功修改了一件商品！');"
'response.write "location.href='"&url&"?page="&page&"';"
response.write "location.href='manage_product2.asp';"
response.write "</script>"
response.end
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
</head>
<script language="JavaScript" type="text/JavaScript">
function JM_wu(ob){
 ob.style.display="none";
}
function JM_you(ob){
 ob.style.display="";
}
function uppic(model,frmname) {
 id=document.prodtable.ProdId.value;
 if (id=='') {
 alert("请先填写商品编号！");
 document.prodtable.ProdId.focus();
return false;}
window.open("../upload/manage_upload.asp?fuptype=pic&fupname="+id+model+"&frmname="+frmname,"blank_","scrollbars=yes,resizable=no,width=650,height=450")
}

function CheckAll(form) {
 for (var i=0;i<form.elements.length;i++) {
 var e = form.elements[i];
 if (e.name != 'chkall') e.checked = form.chkall.checked; 
 }
 }
</script>
<BODY>
<%
id=Trim(request("id"))
if id="" then
	response.write "<script language='javascript'>"
	response.write "alert('商品不存在，请检查！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end if
set rs = server.createobject("adodb.recordset")
sql = "select * from dk501_product where ProdId='"&id&"'"
rs.open sql,conn,1,1

if rs.eof and rs.bof then
response.write "<script language='javascript'>"
response.write "alert('没有此商品序号，请检查！');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
if not (rs.eof and rs.bof) then 
%>
<TABLE cellSpacing=0 cellPadding=0 width=99% border=0 style="BORDER: #8B979F 1px solid; ">
  <TR>
    <TD height="30" bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000">&nbsp;编辑产品</font></B></font></TD>
  </TR>
</TABLE>
<table height="5">
  <tr>
    <td></td>
  </tr>
</table>
<table width="99%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F">
<form name=prodtable method="post" action='manage_product3.asp'>
<%
'从商品资料中读取分类
LarCode=rs("LarCode")
MidCode=rs("MidCode")

'从分类表中读取所有分类
set rsmid=server.createobject("adodb.recordset") 
sqlmid="select distinct LarCode,LarSeq,MidCode,midSeq from dk501_productclass order by LarSeq,MIdSeq " 
rsmid.Open sqlmid,conn,1,1 
%>
  <TR>
    <TD width=100 height="25" align=center bgcolor=#FFFFFF>商品分类：</td>
    <td height="25" colspan=2 bgcolor=#FFFFFF><%
response.write "<select name='MidCode'>"
if rsmid.bof and rsmid.eof then
response.write "<option selected value='未归类^未归类'>-暂无分类-</option>"
else
 response.write "<option value='未归类^未归类' >-该商品不归类-</option>"
 Do While Not rsmid.eof
a=rsmid("larcode")
 response.write "<option value='"&rsmid("larCode")&"^"&rsmid("MidCode")&"'"
 if MidCode=rsmid("MidCode") and larCode=rsmid("larCode") then response.write " selected"
 response.write ">"&rsmid("LarSeq")&"-"&rsmid("LarCode")&" >>"&rsmid("MidSeq")&"-"&rsmid("MidCode")&"</option>"
 rsmid.movenext
if rsmid.eof then exit do
if rsmid("larcode")<>a then response.write "<option value=''>--------------------------</option>"
 loop
end if
rsmid.close
set rsmid=nothing
%></TD>
  </TR>
  <TR>
    <TD height="25" align=center bgcolor=#FFFFFF>商品编号：</td>
    <td height="25" colspan="2" bgcolor=#FFFFFF><input name="ProdId" readonly type="text" size="20" maxlength="20" value="<%=rs("ProdId")%>">
      &nbsp;<img src=../Images/folderclosed.gif alt='商品编号不能修改' width="18" height="18"></TD>
  </TR>
  <TR>
    <TD height="25" align=center bgcolor=#FFFFFF>商品名字：</td>
    <td height="25" bgcolor=#FFFFFF colspan="2"><input name="ProdName" type="text" value="<%=rs("ProdName")%>" size="80" maxlength="80">
      &nbsp;<img src=../Images/folderclosed.gif alt='必填，不能为空' width="18" height="18"></TD>
  </TR>
  <TR>
    <TD height="25" align=center bgcolor=#FFFFFF>商品尺码：</td>
    <td height="25" bgcolor=#FFFFFF colspan="2"><input name="prod1" type="text" value="<%=rs("prod1")%>" size="30">
      &nbsp;尺码标注如;M170/L175/XL/180</TD>
  </TR>
  <TR bgcolor="#FFFFFF">
    <TD height="25" align=center>商品属性：</td>
    <td height="25" colspan="2">单件重量：
      <input name="zhongliang" type="text" value="<%=rs("zhongliang")%>" size="6" maxlength="10">
      克 市场价 ：
      <input name="PriceOrigin" type="text" id="PriceOrigin" value="<%=rs("PriceOrigin")%>" size="6" maxlength="10">
      元 最低售价：
      <input name="PriceList" type="text" id="PriceList" value="<%=rs("PriceList")%>" size="6" maxlength="10">
      元&nbsp;</TD>
  </TR>
  <TR bgcolor="#FFFFFF">
    <TD height="25" align=center>会 &nbsp;员 价：</td>
    <td height="25" colspan="2"><%
	      Set rs123 = conn.Execute("select usertype1,usertype2,usertype3,usertype4,usertype5,usertype6 from dk501_bconfig") 
		usertype1=rs123("usertype1")
		usertype2=rs123("usertype2")
		usertype3=rs123("usertype3")
		usertype4=rs123("usertype4")
		usertype5=rs123("usertype5")
		usertype6=rs123("usertype6")
		  	rs123.close
	set rs123=nothing
	%>
      <%if usertype1<>"" then %>
      <font color="#0000FF"><%=usertype1%>：</font>
      <input name="PriceList1" type="text" size="6" maxlength="10" value="<%=rs("PriceList1")%>" />
      元&nbsp;
      <%end if %>
      <%if usertype2<>"" then %>
      <font color="#0000FF"><%=usertype2%>：</font>
      <input name="PriceList2" type="text" value="<%=rs("PriceList2")%>" size="6" maxlength="10" />
      元&nbsp;
      <%end if %>
      <%if usertype3<>"" then %>
      <font color="#0000FF"><%=usertype3%>：</font>
      <input name="PriceList3" type="text" value="<%=rs("PriceList3")%>" size="6" maxlength="10" />
      元&nbsp;
      <%end if %>
      <%if usertype4<>"" then %>
      <font color="#0000FF"><%=usertype4%>：</font>
      <input name="PriceList4" type="text" value="<%=rs("PriceList4")%>" size="6" maxlength="10" />
      元&nbsp;
      <%else%>
      <input name="PriceList4"  type="hidden" value="0" size="6" maxlength="10" />
      <%end if %>
      <%if usertype5<>"" then %>
      <font color="#0000FF"><%=usertype5%>：</font>
      <input name="PriceList5" type="text" value="<%=rs("PriceList5")%>" size="6" maxlength="10" />
      元&nbsp;
      <%else%>
      <input name="PriceList5"  type="hidden" value="0" size="6" maxlength="10" />
      <%end if %>
      <%if usertype6<>"" then %>
      <font color="#0000FF"><%=usertype6%>：</font>
      <input name="PriceList6" type="text" value="<%=rs("PriceList6")%>" size="6" maxlength="10" />
      元&nbsp;
      <%else%>
      <input name="PriceList6"  type="hidden" value="0" size="6" maxlength="10" />
      <%end if %>    </TD>
  </TR>
  <TR bgcolor="#FFFFFF">
    <TD height="25" align=center>获得积分：</td>
    <td height="25" bgcolor="#EFEFEF" colspan="2"><input name="jifenzhi" type="text" value="<%=rs("jifenzhi")%>" size="6" maxlength="10">
      <span class="800_left_red">产品货号：
      <input name="Product_number" type="text" id="Product_number" value="<%=rs("Product_number")%>" size="10" maxlength="10">
      产品价格：
      <input name="Product_Price" type="text" id="Product_Price" value="<%=rs("Product_Price")%>" size="6" maxlength="10" onBlur="javascript:a();" onKeyUp="value=value.replace(/[^\d]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))">
      产品厂家：
      <select name="Product_Factory" id="Product_Factory" >
	  <%				 
	Set rs_mall = conn.Execute("select UserId from dk501_user where mall=2")
	if rs_mall.bof and rs_mall.eof then 
		response.Write "&nbsp;没有任何栏目"
	else	
		do while not rs_mall.eof
	%>
	<option value="<%=rs_mall("UserId")%>" <%if rs_mall("UserId")=rs("Product_Factory") then%>selected<%end if%>>== <%=rs_mall("UserId")%> == </option> 
   <%
		rs_mall.movenext
		loop
	end if 
	rs_mall.close
	Set rs_mall=Nothing
	%>
      </select>    </TD>
  </TR>
  <TR bgcolor="#FFFFFF">
    <TD height="25" align=center>备货状态：</td>
    <td height="25" colspan="2">
	<input type="radio" name="Quantity" value="1" <%if rs("Quantity")="1" then%>checked<%end if%>>有货&nbsp; 
	  <input type="radio" name="Quantity" value="-1" <%if rs("Quantity")="-1" then%>checked<%end if%>>缺码&nbsp; 
	  <input type="radio" name="Quantity" value="0" <%if rs("Quantity")="0" then%>checked<%end if%>>断货&nbsp;
	  <input type="radio" name="Quantity" value="2" <%if rs("Quantity")="2" then%>checked<%end if%>>大码&nbsp;</TD>
  </TR>
  <TR bgcolor="#FFFFFF">
    <TD height="25" align=center>商品略图：</td>
    <td height="25" colspan="2"><input name="ImgPrev" type="text" value="<%=rs("ImgPrev")%>" size="60" maxlength="100">
      &nbsp;
      <INPUT TYPE="button" value="浏览" onClick="javascript:uppic('_160','ImgPrev');">
      &nbsp;&nbsp;</TD>
  </TR>
  <TR bgcolor="#FFFFFF">
    <TD height="25" align=center>商品大图：</td>
    <td height="25" colspan="2"><input name="ImgPrev2" type="text" value="<%=rs("ImgPrev2")%>" size="60" maxlength="100">
      &nbsp;
      <INPUT name="button" TYPE="button" onClick="javascript:uppic('_300','ImgPrev2');" value="浏览">
      &nbsp;&nbsp;</TD>
  </TR>
  <TR bgcolor="#FFFFFF">
    <TD height="25" align=center>商品细节：</td>
    <td height="25" colspan="2"><input name="MemoSpec" type="text" id="MemoSpec" value="<%=rs("MemoSpec")%>" size="60" maxlength="100">
      &nbsp;
      <INPUT name="button" TYPE="button" onClick="javascript:uppic('_pic','MemoSpec');" value="浏览">
      &nbsp;&nbsp;</TD>
  </TR>
    <TR bgcolor="#FFFFFF">
    <TD height="25" align=center>细节广告图：</td>
    <td height="25" colspan="2"><input name="MemoSpec_AD" type="text" id="MemoSpec_AD" value="<%=rs("MemoSpec_AD")%>" size="60" maxlength="100">
      &nbsp;
      <INPUT name="button" TYPE="button" onClick="javascript:uppic('_pic','MemoSpec_AD');" value="浏览" disabled>
      &nbsp;&nbsp;</TD>
  </TR>
  <TR bgcolor="#FFFFFF">
    <TD height="25" align=center>PS源文件：</td>
    <td height="25" colspan="2"><input name="MemoSpec_PSD" type="text" id="MemoSpec_PSD" value="<%=rs("MemoSpec_PSD")%>" size="60" maxlength="100">
      &nbsp;
      <INPUT name="button" TYPE="button" onClick="javascript:uppic('_pic','MemoSpec_PSD');" value="浏览" disabled>
      &nbsp;&nbsp;</TD>
  </TR>
  <TR bgcolor="#FFFFFF">
    <TD height="25" align=center valign="middle">商品简介：<br>
      (200字内)</td>
    <TD height="25" colspan="2" align=left valign="middle"><textarea name="ProdDisc" cols="80" rows="4" id="ProdDisc"><%=rs("ProdDisc")%></textarea></td>
  </TR>
  <TR bgcolor="#FFFFFF">
    <TD width=73 height="25" align=center>是否推荐</td>
    <td height="25" colspan="2"><input type="radio" name="Remark" value="0" <%if rs("Remark")=0 then%>checked<%end if%>>否 
        <input type="radio" name="Remark" value="1" <%if rs("Remark")=1 then%>checked<%end if%>>
        是 &nbsp;<img src=../Images/folderclosed.gif alt='选择“是”将设为推荐商品' width="18" height="18"> 是否提升：
        <input type="radio" name="up" value="0" checked="checked" />
        否&nbsp;
        <input type="radio" name="up" value="1">
是 &nbsp;<img src="../Images/folderclosed.gif" alt='若选择“是”，将实现：&lt;br&gt;前台：此商品将成为首页第一个商品&lt;br&gt;后台：出现在列表顶端' width="18" height="18" /> 浏览次数:此商品已被浏览
<input type="text" name="ClickTimes" value="<%=rs("ClickTimes")%>" size="5" maxlength="5" />
次 &nbsp;<img src="../Images/folderclosed.gif" alt="小技巧：&lt;br&gt;要让此商品出现在热卖排行的前面位置&lt;br&gt;只需将此数字改得尽量大即可&lt;br&gt;&lt;font color=red&gt;最大的排在第一位置&lt;/font&gt;" width="18" height="18" /></TD>
  </TR>
  <TR bgcolor="#FFFFFF">
    <TD height="25" colspan=3><INPUT TYPE="hidden" name=modify value=edit>
      <input type="submit" name="action" value="添加完成"></TD>
  </TR>
  </form>
</TABLE>
<%
rs.close
set rs=nothing
end if
%>
</BODY>
</html>
<% call CloseConn() %>
