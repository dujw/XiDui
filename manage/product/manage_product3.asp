<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<%
'�޸���Ʒ����
if request("modify")="edit" then
	if request("midcode")="" then
		response.write "<script language='javascript'>"
		response.write "alert('��û��ѡ����Ч����Ʒ���࣬������ȷ����������һҳ��');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
	response.end
	end if
if request("ProdId")=""  or request("Prodname")="" or request("PriceList")="" or request("PriceOrigin")="" then
	response.write "<script language='javascript'>"
	response.write "alert('��Ʒ������д��������������ȷ����������һҳ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end if

if len(request("ProdDisc"))>200 then
response.write "<script language='javascript'>"
response.write "alert('����д����Ʒ�������̫���ˣ�������ȷ����������һҳ��');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if

id=Trim(request("ProdId"))
if id="" then
	response.write "<script language='javascript'>"
	response.write "alert('��Ʒ�����ڣ����飡');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end if
set rs=server.createobject("adodb.recordset")
sql = "select * from dk501_product where ProdId='"&id&"'"
rs.open sql,conn,1,3
if not (rs.eof and rs.bof) then
	rs("ProdId")=Trim(rs("ProdId"))
	rs("ProdName")=Trim(request.form("ProdName"))'��Ʒ����
	rs("prod1")=Trim(request.form("prod1"))'�Զ���һ	
	rs("zhongliang")=Trim(request.form("zhongliang"))
	rs("PriceOrigin")=Trim(request.form("PriceOrigin"))'ԭ��
	rs("PriceList")=Trim(request.form("PriceList"))'�ּ�
	code=split(request.form("Midcode"),"^")
	rs("LarCode")=code(0)'����
	rs("MidCode")=code(1)'����
	rs("PriceList1")=request.form("PriceList1")'�ּ�
	rs("PriceList2")=request.form("PriceList2")'�ּ�
	rs("PriceList3")=request.form("PriceList3")'�ּ�
	rs("PriceList4")=request.form("PriceList4")'�ּ�
	rs("PriceList5")=request.form("PriceList5")'�ּ�
	rs("PriceList6")=request.form("PriceList6")'�ּ�
	rs("jifenzhi")=request.form("jifenzhi")'����
	rs("Product_Price")=Int(request.form("Product_Price"))'��Ʒ�۸�
	rs("Product_number")=Trim(request.form("Product_number"))'��Ʒ����
	rs("Product_Factory")=request.form("Product_Factory")'��Ʒ����
	rs("ProdDisc")=Trim(request.form("ProdDisc"))'���
	rs("ImgPrev")=Trim(request.form("ImgPrev"))'��ƷͼƬ
	rs("ImgPrev2")=Trim(request.form("ImgPrev2"))'��ƷͼƬ
	rs("MemoSpec")=Trim(request.form("MemoSpec"))'��Ʒ����
	
	MemoSpec_AD=request.form("MemoSpec_AD")	'ϸ�ڹ��ͼ�ļ�·��
	if MemoSpec_AD="" then
		MemoSpec_AD=null
	end if
	rs("MemoSpec_AD")=MemoSpec_AD
	
	MemoSpec_PSD=request.form("MemoSpec_PSD")	'PSDԴ�ļ�·��
	if MemoSpec_PSD="" then
		MemoSpec_PSD=null
	end if
	rs("MemoSpec_PSD")=MemoSpec_PSD
	
	rs("Remark")=request.form("Remark")'�Ƿ��Ƽ�
	rs("Quantity")=request.form("Quantity")
	if request.form("up")="1" then rs("TJDate")=now()'�Ƿ�����
rs.update
rs.close
set rs=nothing
end if
response.write "<script language='javascript'>"
response.write "alert('���ѳɹ��޸���һ����Ʒ��');"
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
 alert("������д��Ʒ��ţ�");
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
	response.write "alert('��Ʒ�����ڣ����飡');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end if
set rs = server.createobject("adodb.recordset")
sql = "select * from dk501_product where ProdId='"&id&"'"
rs.open sql,conn,1,1

if rs.eof and rs.bof then
response.write "<script language='javascript'>"
response.write "alert('û�д���Ʒ��ţ����飡');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
if not (rs.eof and rs.bof) then 
%>
<TABLE cellSpacing=0 cellPadding=0 width=99% border=0 style="BORDER: #8B979F 1px solid; ">
  <TR>
    <TD height="30" bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000">&nbsp;�༭��Ʒ</font></B></font></TD>
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
'����Ʒ�����ж�ȡ����
LarCode=rs("LarCode")
MidCode=rs("MidCode")

'�ӷ�����ж�ȡ���з���
set rsmid=server.createobject("adodb.recordset") 
sqlmid="select distinct LarCode,LarSeq,MidCode,midSeq from dk501_productclass order by LarSeq,MIdSeq " 
rsmid.Open sqlmid,conn,1,1 
%>
  <TR>
    <TD width=100 height="25" align=center bgcolor=#FFFFFF>��Ʒ���ࣺ</td>
    <td height="25" colspan=2 bgcolor=#FFFFFF><%
response.write "<select name='MidCode'>"
if rsmid.bof and rsmid.eof then
response.write "<option selected value='δ����^δ����'>-���޷���-</option>"
else
 response.write "<option value='δ����^δ����' >-����Ʒ������-</option>"
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
    <TD height="25" align=center bgcolor=#FFFFFF>��Ʒ��ţ�</td>
    <td height="25" colspan="2" bgcolor=#FFFFFF><input name="ProdId" readonly type="text" size="20" maxlength="20" value="<%=rs("ProdId")%>">
      &nbsp;<img src=../Images/folderclosed.gif alt='��Ʒ��Ų����޸�' width="18" height="18"></TD>
  </TR>
  <TR>
    <TD height="25" align=center bgcolor=#FFFFFF>��Ʒ���֣�</td>
    <td height="25" bgcolor=#FFFFFF colspan="2"><input name="ProdName" type="text" value="<%=rs("ProdName")%>" size="80" maxlength="80">
      &nbsp;<img src=../Images/folderclosed.gif alt='�������Ϊ��' width="18" height="18"></TD>
  </TR>
  <TR>
    <TD height="25" align=center bgcolor=#FFFFFF>��Ʒ���룺</td>
    <td height="25" bgcolor=#FFFFFF colspan="2"><input name="prod1" type="text" value="<%=rs("prod1")%>" size="30">
      &nbsp;�����ע��;M170/L175/XL/180</TD>
  </TR>
  <TR bgcolor="#FFFFFF">
    <TD height="25" align=center>��Ʒ���ԣ�</td>
    <td height="25" colspan="2">����������
      <input name="zhongliang" type="text" value="<%=rs("zhongliang")%>" size="6" maxlength="10">
      �� �г��� ��
      <input name="PriceOrigin" type="text" id="PriceOrigin" value="<%=rs("PriceOrigin")%>" size="6" maxlength="10">
      Ԫ ����ۼۣ�
      <input name="PriceList" type="text" id="PriceList" value="<%=rs("PriceList")%>" size="6" maxlength="10">
      Ԫ&nbsp;</TD>
  </TR>
  <TR bgcolor="#FFFFFF">
    <TD height="25" align=center>�� &nbsp;Ա �ۣ�</td>
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
      <font color="#0000FF"><%=usertype1%>��</font>
      <input name="PriceList1" type="text" size="6" maxlength="10" value="<%=rs("PriceList1")%>" />
      Ԫ&nbsp;
      <%end if %>
      <%if usertype2<>"" then %>
      <font color="#0000FF"><%=usertype2%>��</font>
      <input name="PriceList2" type="text" value="<%=rs("PriceList2")%>" size="6" maxlength="10" />
      Ԫ&nbsp;
      <%end if %>
      <%if usertype3<>"" then %>
      <font color="#0000FF"><%=usertype3%>��</font>
      <input name="PriceList3" type="text" value="<%=rs("PriceList3")%>" size="6" maxlength="10" />
      Ԫ&nbsp;
      <%end if %>
      <%if usertype4<>"" then %>
      <font color="#0000FF"><%=usertype4%>��</font>
      <input name="PriceList4" type="text" value="<%=rs("PriceList4")%>" size="6" maxlength="10" />
      Ԫ&nbsp;
      <%else%>
      <input name="PriceList4"  type="hidden" value="0" size="6" maxlength="10" />
      <%end if %>
      <%if usertype5<>"" then %>
      <font color="#0000FF"><%=usertype5%>��</font>
      <input name="PriceList5" type="text" value="<%=rs("PriceList5")%>" size="6" maxlength="10" />
      Ԫ&nbsp;
      <%else%>
      <input name="PriceList5"  type="hidden" value="0" size="6" maxlength="10" />
      <%end if %>
      <%if usertype6<>"" then %>
      <font color="#0000FF"><%=usertype6%>��</font>
      <input name="PriceList6" type="text" value="<%=rs("PriceList6")%>" size="6" maxlength="10" />
      Ԫ&nbsp;
      <%else%>
      <input name="PriceList6"  type="hidden" value="0" size="6" maxlength="10" />
      <%end if %>    </TD>
  </TR>
  <TR bgcolor="#FFFFFF">
    <TD height="25" align=center>��û��֣�</td>
    <td height="25" bgcolor="#EFEFEF" colspan="2"><input name="jifenzhi" type="text" value="<%=rs("jifenzhi")%>" size="6" maxlength="10">
      <span class="800_left_red">��Ʒ���ţ�
      <input name="Product_number" type="text" id="Product_number" value="<%=rs("Product_number")%>" size="10" maxlength="10">
      ��Ʒ�۸�
      <input name="Product_Price" type="text" id="Product_Price" value="<%=rs("Product_Price")%>" size="6" maxlength="10" onBlur="javascript:a();" onKeyUp="value=value.replace(/[^\d]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))">
      ��Ʒ���ң�
      <select name="Product_Factory" id="Product_Factory" >
	  <%				 
	Set rs_mall = conn.Execute("select UserId from dk501_user where mall=2")
	if rs_mall.bof and rs_mall.eof then 
		response.Write "&nbsp;û���κ���Ŀ"
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
    <TD height="25" align=center>����״̬��</td>
    <td height="25" colspan="2">
	<input type="radio" name="Quantity" value="1" <%if rs("Quantity")="1" then%>checked<%end if%>>�л�&nbsp; 
	  <input type="radio" name="Quantity" value="-1" <%if rs("Quantity")="-1" then%>checked<%end if%>>ȱ��&nbsp; 
	  <input type="radio" name="Quantity" value="0" <%if rs("Quantity")="0" then%>checked<%end if%>>�ϻ�&nbsp;
	  <input type="radio" name="Quantity" value="2" <%if rs("Quantity")="2" then%>checked<%end if%>>����&nbsp;</TD>
  </TR>
  <TR bgcolor="#FFFFFF">
    <TD height="25" align=center>��Ʒ��ͼ��</td>
    <td height="25" colspan="2"><input name="ImgPrev" type="text" value="<%=rs("ImgPrev")%>" size="60" maxlength="100">
      &nbsp;
      <INPUT TYPE="button" value="���" onClick="javascript:uppic('_160','ImgPrev');">
      &nbsp;&nbsp;</TD>
  </TR>
  <TR bgcolor="#FFFFFF">
    <TD height="25" align=center>��Ʒ��ͼ��</td>
    <td height="25" colspan="2"><input name="ImgPrev2" type="text" value="<%=rs("ImgPrev2")%>" size="60" maxlength="100">
      &nbsp;
      <INPUT name="button" TYPE="button" onClick="javascript:uppic('_300','ImgPrev2');" value="���">
      &nbsp;&nbsp;</TD>
  </TR>
  <TR bgcolor="#FFFFFF">
    <TD height="25" align=center>��Ʒϸ�ڣ�</td>
    <td height="25" colspan="2"><input name="MemoSpec" type="text" id="MemoSpec" value="<%=rs("MemoSpec")%>" size="60" maxlength="100">
      &nbsp;
      <INPUT name="button" TYPE="button" onClick="javascript:uppic('_pic','MemoSpec');" value="���">
      &nbsp;&nbsp;</TD>
  </TR>
    <TR bgcolor="#FFFFFF">
    <TD height="25" align=center>ϸ�ڹ��ͼ��</td>
    <td height="25" colspan="2"><input name="MemoSpec_AD" type="text" id="MemoSpec_AD" value="<%=rs("MemoSpec_AD")%>" size="60" maxlength="100">
      &nbsp;
      <INPUT name="button" TYPE="button" onClick="javascript:uppic('_pic','MemoSpec_AD');" value="���" disabled>
      &nbsp;&nbsp;</TD>
  </TR>
  <TR bgcolor="#FFFFFF">
    <TD height="25" align=center>PSԴ�ļ���</td>
    <td height="25" colspan="2"><input name="MemoSpec_PSD" type="text" id="MemoSpec_PSD" value="<%=rs("MemoSpec_PSD")%>" size="60" maxlength="100">
      &nbsp;
      <INPUT name="button" TYPE="button" onClick="javascript:uppic('_pic','MemoSpec_PSD');" value="���" disabled>
      &nbsp;&nbsp;</TD>
  </TR>
  <TR bgcolor="#FFFFFF">
    <TD height="25" align=center valign="middle">��Ʒ��飺<br>
      (200����)</td>
    <TD height="25" colspan="2" align=left valign="middle"><textarea name="ProdDisc" cols="80" rows="4" id="ProdDisc"><%=rs("ProdDisc")%></textarea></td>
  </TR>
  <TR bgcolor="#FFFFFF">
    <TD width=73 height="25" align=center>�Ƿ��Ƽ�</td>
    <td height="25" colspan="2"><input type="radio" name="Remark" value="0" <%if rs("Remark")=0 then%>checked<%end if%>>�� 
        <input type="radio" name="Remark" value="1" <%if rs("Remark")=1 then%>checked<%end if%>>
        �� &nbsp;<img src=../Images/folderclosed.gif alt='ѡ���ǡ�����Ϊ�Ƽ���Ʒ' width="18" height="18"> �Ƿ�������
        <input type="radio" name="up" value="0" checked="checked" />
        ��&nbsp;
        <input type="radio" name="up" value="1">
�� &nbsp;<img src="../Images/folderclosed.gif" alt='��ѡ���ǡ�����ʵ�֣�&lt;br&gt;ǰ̨������Ʒ����Ϊ��ҳ��һ����Ʒ&lt;br&gt;��̨���������б���' width="18" height="18" /> �������:����Ʒ�ѱ����
<input type="text" name="ClickTimes" value="<%=rs("ClickTimes")%>" size="5" maxlength="5" />
�� &nbsp;<img src="../Images/folderclosed.gif" alt="С���ɣ�&lt;br&gt;Ҫ�ô���Ʒ�������������е�ǰ��λ��&lt;br&gt;ֻ�轫�����ָĵþ����󼴿�&lt;br&gt;&lt;font color=red&gt;�������ڵ�һλ��&lt;/font&gt;" width="18" height="18" /></TD>
  </TR>
  <TR bgcolor="#FFFFFF">
    <TD height="25" colspan=3><INPUT TYPE="hidden" name=modify value=edit>
      <input type="submit" name="action" value="������"></TD>
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
