<%@ language="VBScript" %>
<!-- #include file="inc/common.asp" -->
<html>
<head>
  <title>
<%="Hello " & Request("name")%>
</title>
</head>
<body>
  <h1>Welcome, <%=Server.HTMLEncode(Request("name"))%></h1>
  <a href="/user.asp?id=<%=Request("id")%>">Profile</a>
<%If Request("admin") = "1" Then%>
<p>Admin panel: <a href="/admin/">Open</a></p>
<%End If%>
</body>
</html>