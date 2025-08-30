<%@ language="VBScript" %>
<html>
<body>
  <a href="/user.asp?id=<%= Request("id") %>&name=<%= Server.URLEncode(Request("name")) %>">Profile</a>
  <img src="/img/<%= Request("img") %>.png" alt="<%= Request("alt") %>">
  <form action="/save.asp" method="post">
    <input type="text" name="q" value="<%= Replace(Request("q"), '"', '&quot;') %>">
    <input type="submit" value="Go">
  </form>
</body>
<% If Request("admin") = "1" Then %>
<p>Welcome admin</p>
<% End If %>
</html>