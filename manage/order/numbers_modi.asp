<%
if request.cookies("buyok")("admin")="" then response.write "<meta http-equiv='refresh' content='0;URL=../admin.asp'>"
%>

<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<%
sub checkmanage(str)
Set mrs = conn.Execute("select * from manage where username='"&request.cookies("buyok")("admin")&"'")
if not (mrs.bof and mrs.eof) then
manage=mrs("manage")
if instr(manage,str)<=0 then
response.write "<script language='javascript'>"
response.write "alert('���棺��û�д��������Ȩ�ޣ�');"
response.write "location.href='quit.asp';"
response.write "</script>"
response.end
else
session("buyok_admin_login")=0
end if
else 
response.write "<script language='javascript'>"
response.write "alert('û�е�½������ִ�д˲�����');"
response.write "location.href='quit.asp';"
response.write "</script>"
response.end
end if
set mrs=nothing
end sub
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
<HEAD>
  <title>���ﵥ��ϵͳ</title>  
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" />
  <script language="javascript" src="js/area.js"></script>
</HEAD>
<BODY bgcolor="#D9EAFC">
<%
dim count
dim action
action=trim(request("action"))
%>

<%
if request("no")="modi" then
newsid=request("newsid")
	numbers=Trim(request("numbers"))  '���ﵥ��	
	If Trim(Request.form("numbers"))="" Then
		Response.Write("<script language=""JavaScript"">alert(""������������ﵥ�ţ���"");history.go(-1);</script>")
		response.end
	end if
	Company=Trim(request("Company"))        '��ݷ�ʽ
	Consignee=Trim(request("Consignee"))	'�ջ����˵�����
	photo=Trim(request("photo"))'�ջ����˵ĵ绰	
	product=request.form("product") '�Ƿ��յ�����
	more=request.form("more")       '�Ƿ��Ѿ�����
	guishu=request.form("guishu") 
		if more="0" then
			Reason=""
		else
			Reason=Trim(request.form("Reason"))
				If Trim(Request.form("Reason"))="" Then
				Response.Write("<script language=""JavaScript"">alert(""����������û��������ϸ���ܣ�"");history.go(-1);</script>")
				response.end
			end if
		end if 	

set rs=server.createobject("adodb.recordset")
sql="select * from translist where id="&newsid
rs.open sql,conn,1,3
			rs("numbers")=numbers
			rs("Company")=Company
			rs("Consignee")=Consignee
			rs("photo")=photo
			rs("product")=product
			rs("more")=more
			rs("Reason")=Reason	
			rs("guishu")=guishu	
				if more = 1 then
					rs("time")=now()&" | "&request.cookies("buyok")("admin")
				end if	
rs.update
rs.close
set rs=nothing
response.write "<script language='javascript'>" & chr(13)
		response.write "alert('�����޸ĳɹ���');" & Chr(13)
		response.write "window.document.location.href='numberslist.asp';"&Chr(13)
		response.write "</script>" & Chr(13)
Response.End
end if%>
<table height="5"><tr><td></td></tr></table>
<TABLE cellSpacing=0 cellPadding=0 width=750 border=0 style="BORDER-RIGHT: #b1bfee 1px solid; BORDER-TOP: #b1bfee 1px solid; BORDER-LEFT: #b1bfee 1px solid; BORDER-BOTTOM: #b1bfee 1px solid" 
>
  
    <TR> 
      <TD bgColor=#f4f5fd height=30>��<B><font color="#FF0000">&nbsp;���˵������޸�</font></B></TD>
    </TR>
  
</TABLE>
<% 
		newsid=request("id")
		Set rso=Server.CreateObject("ADODB.RecordSet") 
		sql="select * from translist where  id="&newsid
		rso.Open sql,conn,1,1
			if rso.eof and rso.bof then
				response.Write("û�м�¼")				
			else
			Company=Trim(rso("Company"))
	%> 
<table height="5"><tr><td></td></tr></table>
   <table width="750"  border="0" cellpadding="5" cellspacing="1" bgcolor="#b1bfee">
     <form name="addNEWS" method="post" action="?no=modi">
       <TR bgcolor="#FFFFFF">
         <TD width="12%" nowrap>���˵��ţ�             </TD>
         <TD colspan="2"><input name="numbers" type="text" id="numbers" value="<%=rso("numbers")%>" maxlength="20">
         ���磺2007282627 (<span class="800_hzp_red">û�е���ѡ���������</span>)������λ   
		       
           <select name="Company" id="Company">
             <option value="��ͨ" <% if Company = "��ͨ" then response.Write "selected" end if %>>== ��ͨ ==</option>
             <option value="Բͨ" <% if Company = "Բͨ" then response.Write "selected" end if %>>== Բͨ ==</option>
             <option value="˳��" <% if Company = "˳��" then response.Write "selected" end if %>>== ˳�� ==</option>
             <option value="EMS"  <% if Company = "EMS" then response.Write "selected" end if %>>== EMS ==</option>
			 <option value="��ͨ" <% if Company = "��ͨ" then response.Write "selected" end if %>>== ��ͨ ==</option>
			 <option value="����" <% if Company = "����" then response.Write "selected" end if %>>== ���� ==</option>
			 <option value="�г�" <% if Company = "�г�" then response.Write "selected" end if %>>== �г� ==</option>
			 <option value="�ϴ�" <% if Company = "�ϴ�" then response.Write "selected" end if %>>== �ϴ� ==</option>
			 <option value="����" <% if Company = "����" then response.Write "selected" end if %>>== ���� ==</option>
           </select>         </TD>
       </TR> 
       <TR bgcolor="#FFFFFF">
         <TD nowrap>�ʻ������֣�           </TD>
         <TD><input name="photo" type="text" id="photo" value="<%=rso("photo")%>" size="50"></TD>
          <TD width="48%">����ID��<b><font color="#FF0000"><a href="../user/manage_user_edit.asp?action=useredit&id=<%=rso("DU_UserName")%>"><%=rso("DU_UserName")%></a></font></b>
          </TD>
       </TR>
	   <TR bgcolor="#FFFFFF">
         <TD nowrap>����������         </TD>
         <TD colspan="2"><textarea name="Consignee" cols="50" rows="4" id="Consignee"><%=rso("Consignee")%></textarea></TD>
       </TR>
       <TR bgcolor="#FFFFFF">
         <TD nowrap>�˻����飺         </TD>
         <TD colspan="2"><textarea name="Consignee" cols="50" rows="4" id="Consignee" disabled><%=rso("DU_Note")%></textarea></TD>
       </TR>
       <TR bgcolor="#FFFFFF">
         <TD nowrap>���������</TD>
         <TD colspan="2"><input name="guishu"  type="radio" value=1 <%if rso("guishu")=1 then Response.Write "checked"%>  />������</TD>
       </TR>
	   <TR bgcolor="#FFFFFF">
         <TD nowrap>�Ƿ��յ����</TD>
         <TD colspan="2"><input name="product"  type="radio" value=False <%if rso("product")=False then Response.Write "checked"%>  />
��
  <input name="product"  type="radio"  value=True <%if rso("product")=True then Response.Write "checked"%>/>
��</TD>
       </TR>
	   <TR bgcolor="#FFFFFF">
         <TD nowrap>�Ƿ��Ѿ�����
           <script language=javascript>function show(obj)
			{
				if(document.getElementById("r1").checked)
					{
					document.getElementById("PhotoUrl").style.display=''
					}
				else
					{
					document.getElementById("PhotoUrl").style.display='none'
					}
			}
			</script></TD>
         <TD colspan="2"><input name="more"  type="radio" onClick="show()" value="0" <%if rso("more")=False  then Response.Write "checked"%>/>
��
  <input name="more"  type="radio" onClick="show()" value="1" id="r1" <%if rso("more")=True  then Response.Write "checked"%>/><%=rso("more")%>
��
<script language=javascript>
			</script></TD>
       </TR>
	   <tr bgcolor="#ECF5FF"  id="PhotoUrl"  <%if rso("more")="False" then %>style="display:none;"<% end if %>>
         <TD nowrap>���������</TD>
         <TD width="40%"><textarea name="Reason" cols="50" rows="5" id="Reason"><%=rso("Reason")%></textarea></TD>
	     <TD width="48%"><font color="#FF0000">============ģ��====<br>
	     </font><font color="#FF0000">����:  ����:  ������:<br>
�������ϵ����QQ������<br>
	     ====================</font></TD>
	   </TR>
	   <tr bgcolor="#ECF5FF">
         <TD nowrap>�յ�����ʱ��</TD>
         <TD colspan="2"><%=rso("addtime")%></TD>
       </tr>
	   <tr bgcolor="#ECF5FF">
         <TD nowrap>�������ʱ��</TD>
         <TD colspan="2"><%=rso("time")%></TD>
       </tr>	  

       <TR bgcolor="#FFFFFF">
         <TD colspan="3">		
		 <input type="submit" name="submit" value="�޸����">
		 <input type="hidden" name="newsId" value="<%=newsId%>">
		 <input type="reset" name="Submit2" value="�����޸�"></TD>
       </TR>
     </form>
</TABLE> 
         <% End If
rso.close
set rso=nothing
 %>
</body>
</HTML>
