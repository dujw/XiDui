<script language="JavaScript">
function checkform(){
	if (document.loginbox.userid.value.length ==0){
		alert("��ʾ�������������û�����");
		document.loginbox.userid.focus();
		return false;
	}
	if (document.loginbox.password.value.length==0){
		alert("��ʾ���������������룡");
		document.loginbox.password.focus();
		return false;
	}
	return true;
}	
</script>
<%
	Dim check
	Randomize timer
	check=Int((899999)*Rnd +100000)
%>
<table border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td><img src="image/top/user1.jpg" width="195" height="31" /></td>
  </tr>
  <tr>
    <td height="87" align="center" background="image/top/user2.jpg">
    <% if request.cookies("buyok")("userid")="" then %>	
      <table width="180" border="0" align="center" cellpadding="0" cellspacing="0">
      <form method="post" name="loginbox" action="/login/ckeckuser.asp" onSubmit="return checkform();">
        <tr>
          <td height="26">�û�����
            <input name="userid" type="text" id="userid" size="15" /></td>
        </tr>
        <tr>
          <td height="26">�� �� ��
            <input name="password" type="password" id="password" size="15" />       
			</td>
        </tr>
		<tr>
          <td height="26" align="left">&nbsp;��֤�룺
            <input name="check1" type="text" size="6" maxlength="6" />
            <font color="#C04411"><%=check%></font>
            <input type="hidden" name="check2" value="<%=check%>" />
            <input type="hidden" name="login" value="ok" />
            <input type="hidden" name="cook" value="0" /></td>
        </tr>
        <tr>
          <td height="12" align="center"><input type="submit" name="button" id="button" value=" �� ½ " />
              <input type="reset" name="button2" id="button2" value=" ȡ �� " /></td>
        </tr>
        </form>
      </table>
      <% else %>
          �𾴵� <%=request.cookies("buyok")("userid")%> ���:<br>
      <a href="/user/user_center.asp">�ҵĶ���</a> <a href="/product/check.asp">�ҵĹ��ﳵ</a> <br>
          <a href="shande.asp">�˻�����</a><a href="Opinion.asp"></a> <a href="/login/logout.asp">���˳���</a>
      <% end if %>
    </td>
  </tr>
  <tr>
    <td><img src="image/top/user3.jpg" width="195" height="4" /></td>
  </tr>
</table>