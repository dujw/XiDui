<!--#include file="include/connect.asp" -->
<%
call aspsql()
'�ж��û��Ƿ��½
call checklogin()
Name   = request.cookies("buyok")("userid")
'if request("set")="ok" then call setq()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=sitename%></title>
<meta name="keywords" content="<%=sitedescription%>">
<meta name="description" content="<%=sitekeywords%>">
<link href="image/web_css/index.css" rel="stylesheet" type="text/css" />
</head>
<body>
<%
action=trim(request("action"))
if  Request.QueryString("action")="add" then			
	numbers=Trim(request("numbers"))  
	If Trim(Request.form("numbers"))="" Then
		Response.Write("<script language=""JavaScript"">alert(""������������ﵥ�ţ���"");history.go(-1);</script>")
		response.end
	end if
	
	'guishu=Trim(request("guishu")) 
	guishu=1
'	If Trim(Request.form("guishu"))="" Then
'		Response.Write("<script language=""JavaScript"">alert(""������ѡ����Ĺ��޵أ���"");history.go(-1);<\/script>")
'		response.end
'	end if
	Company=Trim(request("Company"))        '��ݷ�ʽ
	Consignee=Trim(request("Consignee"))	'�˻�����
	photo=Trim(request("photo"))'�ջ����˵�����
	'product=request.form("product") '�Ƿ��յ�����
	product=0 '�Ƿ��յ�����		
	UserName=Name   '�û���������л��������û�ID
	set rs=server.createobject("adodb.recordset")
	sql="select * from translist where numbers='"&numbers&"'"
	rs.open sql,conn,1,3
	if not (rs.bof and rs.EOF) then '�Ƿ��Ѿ�������������
	  response.write  "<script language=javascript>alert('��ʾ�����Ѿ�����ͬ���Ļ�����!');history.go(-1);</script>"
	  response.end
	else
		rs.Addnew
			rs("numbers")=numbers
			rs("Company")=Company
			rs("DU_Note")=Consignee
			rs("photo")=photo
			rs("guishu")=guishu
			rs("product")=product
			rs("DU_UserName")=UserName
			rs("addtime")=now()&" | "&UserName
		rs.update
		rs.close
		set rs=nothing
		conn.close  
		set conn=nothing
		response.write  "<script language=javascript>alert('��ʾ����ӳɹ�����ȴ��˻�!');location.href('user_TuiHuoList.asp');</script>"
		'Response.Redirect "user_index.asp"	
		Response.End
	end if
end if 
%>
<!--#include file="include/web_top.asp" -->
<div id="div_Public">
<table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="260" valign="top"><!--#include file="left.asp"--></td>
    <td valign="top">
	          <table width="690" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="25" bgcolor="E43A23">&nbsp;<span class="input_search">�û��˻�</span></td>
                </tr>
                <tr>
                  <td height="5"></td>
                </tr>
                <tr>
                  <td height="5">
				  <table width="650" align=center>
  <tr>
    <td align="center" valign="middle"><TABLE cellSpacing=0 cellPadding=0 width=690 border=0 style="BORDER-RIGHT: #AFAFAF 1px solid; BORDER-TOP: #AFAFAF 1px solid; BORDER-LEFT: #AFAFAF 1px solid; BORDER-BOTTOM: #AFAFAF 1px solid">
        <TR>
          <TD bgColor=#f4f5fd height=25><font color="#000000">��<b>��д�˻���Ϣ</b></font></TD>
        </TR>
    </TABLE></td>
  </tr>
  <tr>
    <td align="left" valign="middle">	
	<table width="95%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#b1bfee">
     <form name="addNEWS" method="post" action="?action=add">
       <TR bgcolor="#FFFFFF">
         <TD width="12%" nowrap>���˵��ţ�             </TD>
         <TD width="88%"><input name="numbers" type="text" id="numbers" value="" maxlength="20">
         ���磺2007282627
            <select name="Company" id="Company">
             <option value="��ͨ" selected="selected">== ��ͨ ==</option>
             <option value="Բͨ">== Բͨ ==</option>
             <option value="˳��">== ˳�� ==</option>
             <option value="EMS">== EMS ==</option>			 
			 <option value="��ͨ">== ��ͨ ==</option>
			 <option value="����">== ���� ==</option>
			 <option value="�г�">== �г� ==</option>
			 <option value="�ϴ�">== �ϴ� ==</option>
			 <option value="����">== ���� ==</option>
           </select></TD>
       </TR>       
	   <TR bgcolor="#FFFFFF">
         <TD nowrap>�Ļ��˵����֣�           </TD>
         <TD><input name="photo" type="text" id="photo" value="" size="50"> 
           <span class="800_title">����ֻд���� û�оͲ�д��</span></TD>
       </TR>
       <TR bgcolor="#FFFFFF">
         <TD nowrap>�˻����飺         </TD>
         <TD><textarea name="Consignee" cols="50" rows="4" id="Consignee"></textarea></TD>
       </TR>
       
       <TR bgcolor="#FFFFFF">
         <TD colspan="2">
		 <input type="submit" name="submit" value="������">
         <input type="reset" name="Submit2" value="������д">		 </TD>
       </TR>
     </form>
   </TABLE>
    </td>
  </tr>
</table>
				  </td>
                </tr>
              </table>
      </td>
  </tr>
</table>

</div>
<!--#include file="include/web_boot.asp" -->
</body>
</html>
<% call CloseConn() %>
