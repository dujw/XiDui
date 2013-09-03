<%
url=request.ServerVariables("QUERY_STRING")
%>

<html>
<head>
<title>正在打开……</title>
<meta http-equiv="refresh" content="0;URL=<%=url%>">
</head>
<body>
正在打开……
</body>
</html>
