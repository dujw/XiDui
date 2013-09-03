<!--#include file="../include/Check_Sql.asp"-->
<!--#include file="../../include/connect.asp "-->
<!--#include file="../../include/md5.asp "-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
<HEAD><title>货物单号系统</title>  
<link href="../Include/admin_css.css" rel="stylesheet" type="text/css" /></HEAD>
<BODY>
<%
action=trim(request("action"))
%>

<%
if  Request.QueryString("action")="add" then			
	numbers=Trim(request("numbers"))  
	If Trim(Request.form("numbers"))="" Then
		Response.Write("<script language=""JavaScript"">alert(""错误：请输入货物单号！！"");history.go(-1);</script>")
		response.end
	end if
	
	guishu=Trim(request("guishu")) 
	If Trim(Request.form("guishu"))="" Then
		Response.Write("<script language=""JavaScript"">alert(""错误：请选择货的归宿地！！"");history.go(-1);</script>")
		response.end
	end if
	Company=Trim(request("Company"))        '快递方式
	Consignee=Trim(request("Consignee"))	'收货物人的姓名
	photo=Trim(request("photo"))'收货物人的名字
	product=request.form("product") '是否收到货物	
	UserName=Trim(request("username"))   '用户自助添加托货单——用户ID
	set rs=server.createobject("adodb.recordset")
	sql="select * from translist where numbers='"&numbers&"'"
	rs.open sql,conn,1,3
	if not (rs.bof and rs.EOF) then '是否已经发过这条数据
	  response.write  "<script language=javascript>alert('提示：你已经发过同样的货单号!');history.go(-1);</script>"
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
      <TD bgColor=#f4f5fd height=30>　<B><font color="#FF0000">&nbsp;换货单号码</font></B></TD>
    </TR>
  
</TABLE>
<table height="5"><tr><td></td></tr></table>
   <table width="95%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#b1bfee">
     <form name="addNEWS" method="post" action="?action=add">
       <TR bgcolor="#FFFFFF">
         <TD width="12%" nowrap>货运单号：             </TD>
         <TD width="88%" colspan="2"><input name="numbers" type="text" id="numbers" value="" maxlength="20">
         例如：2007282627
            <select name="Company" id="Company">
             <option value="申通" selected="selected">== 申通 ==</option>
             <option value="圆通">== 圆通 ==</option>
             <option value="顺丰">== 顺丰 ==</option>
             <option value="EMS">== EMS ==</option>			 
			 <option value="中通">== 中通 ==</option>
			 <option value="天天">== 天天 ==</option>
			 <option value="中诚">== 中诚 ==</option>
			 <option value="韵达">== 韵达 ==</option>
			 <option value="其它">== 其它 ==</option>
           </select></TD>
       </TR>       
	   <TR bgcolor="#FFFFFF">
         <TD nowrap>寄货人的名字：           </TD>
         <TD><input name="photo" type="text" id="photo" value="" size="50"> 
           <span class="800_title">这里只写名字 没有就不写哈</span></TD>
         <TD width="30%">代理ID：
          <input name="username" type="text" id="username" value="" maxlength="20"></TD>
           
       </TR>
       <TR bgcolor="#FFFFFF">
         <TD nowrap>货物描述：         </TD>
         <TD colspan="2"><textarea name="Consignee" cols="50" rows="4" id="Consignee"></textarea></TD>
       </TR>
	   <TR bgcolor="#FFFFFF">
         <TD nowrap>货物归属：</TD>
         <TD colspan="3">
			 <input name="guishu"  type="radio" value=1 checked /> 裁男网
		 </TD>
       </TR>
       <TR bgcolor="#FFFFFF">
         <TD nowrap>是否收到货物：</TD>
         <TD colspan="2"><input name="product"  type="radio" onClick="show()" value="0" />
否
  <input name="product"  type="radio" onClick="show()" value="1" checked/> 
是</TD>
       </TR>
       
       <TR bgcolor="#FFFFFF">
         <TD colspan="3">
		 <input type="submit" name="submit" value="添加完成">
         <input type="reset" name="Submit2" value="重新填写">		 </TD>
       </TR>
     </form>
   </TABLE> 
</body>
</HTML>
