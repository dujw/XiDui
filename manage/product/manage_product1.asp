
<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<%randomize
	d=right("00"&int(99*rnd()),2)
	yy=right(year(date),4)
	mm=right("00"&month(date),2)
	dd=right("00"&day(date),2)
	riqi=yy & mm & dd
	xiaoshi=right("00"&hour(time),2)
	fenzhong=right("00"&minute(time),2)
	miao=right("00"&second(time),2)
	inBillNo=yy & mm & dd & xiaoshi & fenzhong & miao & d
	inBillNo1=yy&"��" & mm &"��"& dd &"��"& xiaoshi& "ʱ"& fenzhong &"��"
	%>
	
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
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
</head>
<BODY>
<%
if request("add") ="ok" then
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="select * from dk501_product Where ProdId='"&request("ProdId")&"'"
	rs.open sql,conn,3,3
		if not rs.eof then
			response.write "<script language='javascript'>"
			response.write "alert('�����ˣ���Ʒ����ظ������飡');"
			response.write "location.href='javascript:history.go(-1)';"
			response.write "</script>"
			response.end
			rs.close
			set rs=nothing
			response.end
		end if
	
	if request("midcode")="" then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ���û��ѡ����Ч����Ʒ���࣡');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if
	
	if request("ProdId")="" or request("Prodname")="" then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ���Ʒ������д��������');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	'response.end
	end if
	
	if len(request("ProdDisc"))>200 then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ�����д�ļ������̫���ˣ�');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	'response.end
	end if
	
	rs.Addnew
	rs("ProdId")=request.form("ProdId")	'���
	rs("ProdName")=request.form("ProdName")	'��Ʒ����
	rs("prod1")=request.form("prod1")	'�Զ���һ
	rs("zhongliang")=request.form("zhongliang")'�ּ�
	'=============================================
	rs("Product_number")=request.form("Product_number")'��Ʒ����
	rs("Product_Price")=request.form("Product_Price")'��Ʒ�۸�
	rs("Product_Factory")=request.form("Product_Factory")'��Ʒ����
	'=============================================
	
	rs("PriceOrigin")=request.form("PriceOrigin")	'ԭ��
	rs("PriceList")=request.form("PriceList")	'�ּ�
	rs("PriceList1")=request.form("PriceList1")	'�ּ�
	rs("PriceList2")=request.form("PriceList2")	'�ּ�
	rs("PriceList3")=request.form("PriceList3")	'�ּ�
	rs("PriceList4")=request.form("PriceList4")	'�ּ�
	rs("PriceList5")=request.form("PriceList5")	'�ּ�
	rs("PriceList6")=request.form("PriceList6")	'�ּ�
	
	rs("jifenzhi")=request.form("jifenzhi")
	code=split(request.form("Midcode"),"^")
	rs("LarCode")=code(0)	'����
	rs("MidCode")=code(1)	'����
	rs("ProdDisc")=request.form("ProdDisc")	'���
	rs("ImgPrev")=request.form("ImgPrev")	'��ƷͼƬ
	rs("ImgPrev2")=request.form("ImgPrev2")	'��ƷͼƬ
	rs("MemoSpec")=request.form("MemoSpec")	'��Ʒ����
	rs("MemoSpec_AD")=request.form("MemoSpec_AD")	'ϸ�ڹ��ͼ�ļ�·��
	rs("MemoSpec_PSD")=request.form("MemoSpec_PSD")	'PSDԴ�ļ�·��
	rs("Remark")=request.form("Remark")	'�Ƿ��Ƽ�
	rs("online")=2 	'�Ƿ�����
	rs.update
	rs.close
	set rs=nothing
	response.write "<script language='javascript'>"
	response.write "alert('��Ʒ��ӳɹ���������ȷ����������Ʒ�б�');"
	response.write "location.href='manage_product2.asp?show=yes';"
	response.write "</script>"
end if
%>
<TABLE cellSpacing=0 cellPadding=0 width=99% border=0 style="BORDER: #8B979F 1px solid; ">
  <TR>
    <TD height="30" bgColor=#f4f5fd><font color="#000000"><B><font color="#FF0000">&nbsp;��Ӳ�Ʒ</font></B></font></TD>
  </TR>
</TABLE>
<table height="5">
  <tr>
    <td></td>
  </tr>
</table>
<table width="99%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#8B979F">
  <form name=prodtable method="post" action="">
  <%
Dim a,b,c 
set rsmid=server.createobject("adodb.recordset") 
sqlmid="select distinct LarCode,LarSeq,MidCode,midSeq from dk501_productclass order by LarSeq,MIdSeq" 
rsmid.Open sqlmid,conn,1,1 
%>
<TR><TD width=100 align=center bgcolor=#FFFFFF>��Ʒ���ࣺ</td><td bgcolor=#FFFFFF colspan=2>
<%
response.write "<select name='MidCode'>"
response.write"<option value='δ����^δ����' selected>---ѡ����Ʒ��������---</option>"
if rsmid.bof and rsmid.eof then
response.write "<option selected value='δ����^δ����'>-������Ʒ����-</option>"
else
Do While Not rsmid.eof
a = rsmid("larcode")
response.write "<option value='"&rsmid("larCode")&"^"&rsmid("MidCode")&"'>"&rsmid("LarSeq")&"-"&rsmid("LarCode")&" >> "&rsmid("MidSeq")&"-"&rsmid("MidCode")&"</option>"
rsmid.movenext
if rsmid.eof then exit do
if rsmid("larcode")<>a then response.write "<option value=''>--------------------------</option>"
loop
end if
rsmid.close
set rsmid=nothing
%>
</TD></TR>

    <TR>
      <TD height="25" align=center bgcolor=#FFFFFF>��Ʒ��ţ�</td>
      <td height="25" colspan="2" bgcolor=#FFFFFF><input name="ProdId" type="text" size="20" maxlength="20" value="<%=inBillNo%>">&nbsp;<img src=../Images/folderclosed.gif alt='��Ʒ��Ų����޸�' width="18" height="18"></TD>
    </TR>
    <TR>
      <TD height="25" align=center bgcolor=#FFFFFF>��Ʒ���֣�</td>
      <td height="25" bgcolor=#FFFFFF colspan="2"><input name="ProdName" type="text" value="" size="80" maxlength="80">
        &nbsp;<img src=../Images/folderclosed.gif alt='�������Ϊ��' width="18" height="18"></TD>
    </TR>
    <TR>
      <TD height="25" align=center bgcolor=#FFFFFF>��Ʒ���룺</td>
      <td height="25" bgcolor=#FFFFFF colspan="2"><input name="prod1" type="text" size="30">
        &nbsp;�����ע��;M170/L175/XL/180</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="25" align=center>��Ʒ���ԣ�</td>
      <td height="25" colspan="2">����������
        <input name="zhongliang" type="text" value="0" size="6" maxlength="10">
        �� �г��� ��
        <input name="PriceOrigin" type="text" id="PriceOrigin" value="0" size="6" maxlength="10">
        Ԫ ����ۼۣ�
        <input name="PriceList" type="text" id="PriceList" value="0" size="6" maxlength="10">
        Ԫ&nbsp;</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="25" align=center>�� &nbsp;Ա �ۣ�</td>
      <td height="25" colspan="2"><%
	      Set rs = conn.Execute("select usertype1,usertype2,usertype3,usertype4,usertype5,usertype6 from dk501_bconfig") 
	%>
        <%if rs("usertype1")<>"" then %>
        <font color="#0000FF"><%=rs("usertype1")%>��</font>
        <input name="PriceList1" type="text" value="0" size="6" maxlength="10">
        Ԫ&nbsp;
        <%end if %>
        <%if rs("usertype2")<>"" then %>
        <font color="#0000FF"><%=rs("usertype2")%>��</font>
        <input name="PriceList2" type="text" value="0" size="6" maxlength="10">
        Ԫ&nbsp;
        <%end if %>
        <%if rs("usertype3")<>"" then %>
        <font color="#0000FF"><%=rs("usertype3")%>��</font>
        <input name="PriceList3" type="text" value="0" size="6" maxlength="10">
        Ԫ&nbsp;
        <%end if %>
        <%if rs("usertype4")<>"" then %>
        <font color="#0000FF"><%=rs("usertype4")%>��</font>
        <input name="PriceList4" type="text" value="0" size="6" maxlength="10">
        Ԫ&nbsp;
        <%end if %>
        <%if rs("usertype5")<>"" then %>
        <font color="#0000FF"><%=rs("usertype5")%>��</font>
        <input name="PriceList5" type="text" value="0" size="6" maxlength="10">
        Ԫ&nbsp;
        <%end if %>
        <%if rs("usertype6")<>"" then %>
        <font color="#0000FF"><%=rs("usertype6")%>��</font>
        <input name="PriceList6" type="text" value="0" size="6" maxlength="10">
        Ԫ&nbsp;
        <%end if %>
        <%
	rs.close
	set rs=nothing	
	%>
      </TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="25" align=center>��û��֣�</td>
      <td height="25" bgcolor="#EFEFEF" colspan="2"><input name="jifenzhi" type="text" value="0" size="6" maxlength="10">
        <span class="800_left_red">��Ʒ���ţ�
        <input name="Product_number" type="text" id="Product_number" size="10" maxlength="10">
        ��Ʒ�۸�
        <input name="Product_Price" type="text" id="Product_Price" value="0" size="6" maxlength="10" onBlur="javascript:a();" onKeyUp="value=value.replace(/[^\d]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))">
        ��Ʒ���ң�
        <select name="Product_Factory" id="Product_Factory" >
          <%				 
	Set rs_mall = conn.Execute("select UserId from dk501_user where mall=2 ")
	if rs_mall.bof and rs_mall.eof then 
		response.Write "&nbsp;û���κ���Ŀ"
	else	
		do while not rs_mall.eof
	%>
          <option value="<%=rs_mall("UserId")%>">== <%=rs_mall("UserId")%> == </option>
          <%
		rs_mall.movenext
		loop
	end if 
	rs_mall.close
	Set rs_mall=Nothing
	%>
        </select>
      </TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="25" align=center>����״̬��</td>
      <td height="25" colspan="2"><input type="radio" name="Quantity" value="1" checked>
        �л�&nbsp;
        <input type="radio" name="Quantity" value="-1">
        ȱ��
        &nbsp;
        <input type="radio" name="Quantity" value="0">
        �ϻ�
        <input type="radio" name="Quantity" value="2">
        ����
        &nbsp;&nbsp;</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="25" align=center>��Ʒ��ͼ��</td>
      <td height="25" colspan="2"><input name="ImgPrev" type="text" value="UploadFiles/none.gif" size="60" maxlength="100">
        &nbsp;
        <INPUT TYPE="button" value="���" onClick="javascript:uppic('_160','ImgPrev');">
        &nbsp;&nbsp;</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="25" align=center>��Ʒ��ͼ��</td>
      <td height="25" colspan="2"><input name="ImgPrev2" type="text" value="UploadFiles/none.gif" size="60" maxlength="100">
        &nbsp;
        <INPUT name="button" TYPE="button" onClick="javascript:uppic('_300','ImgPrev2');" value="���">
        &nbsp;&nbsp;</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="25" align=center>��Ʒϸ�ڣ�</td>
      <td height="25" colspan="2"><input name="MemoSpec" type="text" id="MemoSpec" value="UploadFiles/none.gif" size="60" maxlength="100">
        &nbsp;
        <INPUT name="button" TYPE="button" onClick="javascript:uppic('_pic','MemoSpec');" value="���">
        &nbsp;&nbsp;</TD>
    </TR>
        <TR bgcolor="#FFFFFF">
      <TD height="25" align=center>ϸ�ڹ��ͼ��</td>
      <td height="25" colspan="2"><input name="MemoSpec_AD" type="text" id="MemoSpec_AD" value="" size="60" maxlength="100">
        &nbsp;
        <INPUT name="button" TYPE="button" onClick="javascript:uppic('_adv','MemoSpec_AD');" value="���" disabled>
        &nbsp;&nbsp;</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="25" align=center>PSԴ�ļ���</td>
      <td height="25" colspan="2"><input name="MemoSpec_PSD" type="text" id="MemoSpec_PSD" value="" size="60" maxlength="100">
        &nbsp;
        <INPUT name="button" TYPE="button" onClick="javascript:uppic('_psd','MemoSpec_PSD');" value="���" disabled>
        &nbsp;&nbsp;</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="25" align=center valign="middle">��Ʒ��飺<br>
        (200����)</td>
      <TD height="25" colspan="2" align=left valign="middle"><textarea name="ProdDisc" cols="80" rows="4" id="ProdDisc"></textarea></td>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD width=73 height="25" align=center>�Ƿ��Ƽ�</td>
      <td height="25" colspan="2"><input type="radio" name="Remark" value="0" checked>
        ��&nbsp;
        <input type="radio" name="Remark" value="1">
        �� &nbsp;<img src=../Images/folderclosed.gif alt='ѡ���ǡ�����Ϊ�Ƽ���Ʒ' width="18" height="18"></TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="25" colspan=3><INPUT TYPE="hidden" name=add value=ok>
        <input type="submit" name="action" value="������"></TD>
    </TR>
  </form>
</TABLE>
</BODY>
</html>
<% call CloseConn() %>
