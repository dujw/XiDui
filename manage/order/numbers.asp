<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<!--#include file="../../include/md5.asp "-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
<HEAD><title>���ﵥ��ϵͳ</title>  
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" /></HEAD>
<BODY>
<%
action=trim(request("action"))
%>

<%
if  Request.QueryString("action")="add" then			
	numbers=Trim(request("numbers"))  
	If Trim(Request.form("numbers"))="" Then
		Response.Write("<script language=""JavaScript"">alert(""������������ﵥ�ţ���"");history.go(-1);</script>")
		response.end
	end if
	
	guishu=Trim(request("guishu")) 
	If Trim(Request.form("guishu"))="" Then
		Response.Write("<script language=""JavaScript"">alert(""������ѡ����Ĺ��޵أ���"");history.go(-1);</script>")
		response.end
	end if
	Company=Trim(request("Company"))        '��ݷ�ʽ
	Consignee=Trim(request("Consignee"))	'�ջ����˵�����
	photo=Trim(request("photo"))'�ջ����˵�����
	product=request.form("product") '�Ƿ��յ�����	
	UserName=Trim(request("username"))   '�û���������л��������û�ID
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
			rs("Consignee")=Consignee
			rs("photo")=photo
			rs("guishu")=guishu
			rs("product")=product
			rs("DU_UserName")=UserName
			rs("addtime")=now()&" | "&request.cookies("buyok")("admin")
		rs.update
		rs.close
		set rs=nothing
		conn.close  
		set conn=nothing
		Response.Redirect "numberslist.asp"	
		Response.End
	end if
end if 
%>
<table height="5"><tr><td></td></tr></table>
<TABLE cellSpacing=0 cellPadding=0 width=95% border=0 style="BORDER-RIGHT: #b1bfee 1px solid; BORDER-TOP: #b1bfee 1px solid; BORDER-LEFT: #b1bfee 1px solid; BORDER-BOTTOM: #b1bfee 1px solid" 
>
  
    <TR> 
      <TD bgColor=#f4f5fd height=30>��<B><font color="#FF0000">&nbsp;����������</font></B></TD>
    </TR>
  
</TABLE>
<table height="5"><tr><td></td></tr></table>
   <table width="95%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#b1bfee">
     <form name="addNEWS" method="post" action="?action=add">
       <TR bgcolor="#FFFFFF">
         <TD width="12%" nowrap>���˵��ţ�             </TD>
         <TD width="88%" colspan="2"><input name="numbers" type="text" id="numbers" value="" maxlength="20">
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
         <TD width="30%">����ID��
          <input name="username" type="text" id="username" value="" maxlength="20"></TD>
           
       </TR>
       <TR bgcolor="#FFFFFF">
         <TD nowrap>����������         </TD>
         <TD colspan="2"><textarea name="Consignee" cols="50" rows="4" id="Consignee"></textarea></TD>
       </TR>
	   <TR bgcolor="#FFFFFF">
         <TD nowrap>���������</TD>
         <TD colspan="3">
			 <input name="guishu"  type="radio" value=1 checked /> ������
		 </TD>
       </TR>
       <TR bgcolor="#FFFFFF">
         <TD nowrap>�Ƿ��յ����</TD>
         <TD colspan="2"><input name="product"  type="radio" onClick="show()" value="0" />
��
  <input name="product"  type="radio" onClick="show()" value="1" checked/> 
��</TD>
       </TR>
       
       <TR bgcolor="#FFFFFF">
         <TD colspan="3">
		 <input type="submit" name="submit" value="������">
         <input type="reset" name="Submit2" value="������д">		 </TD>
       </TR>
     </form>
   </TABLE> 
</body>
</HTML>
