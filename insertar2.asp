<html>
<body>
<%@ LANGUAGE="VBSCRIPT" %>
<%
option explicit
Dim conn,nombre,segundoa,primera,email,mensajes

nombre=Request.Form("nombre")
primera=Request.Form("primera")
segundoa=Request.Form("segundoa")
email=Request.Form("email")
mensajes=Request.Form("mensaje")

set conn=Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Carlos Chavez\Documents\cqrrillo\compras.mdb" 
conn.execute "INSERT INTO formulario(nombre,primera,segundoa,email,mensaje) values('" & nombre & "','" & primera & "','" & segundoa & "','" & email &  "','" & mensajes & "')"
conn.close
set conn=nothing
response.redirect("formulario.html")

%>
</body>
</html>