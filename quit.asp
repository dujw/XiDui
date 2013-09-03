<html>
<head>
<script language='javascript'>
if (top != self)top.location.href = "/manage/login.asp"; 
</script>
<meta http-equiv="refresh" content="0;URL=/manage/login.asp">
</head>
<%
response.cookies("buyok")("admin")=""
session.abandon
%>
</html>