<html>
<body>
<%@ LANGUAGE="VBSCRIPT" %>
<%
option explicit
Dim conn,prod1,prod2,prod3,cant

prod1=Request.Form("prod1")
prod2=Request.Form("prod2")
prod3=Request.Form("prod3")
cant=Request.Form("cantidad")
set conn=Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Carlos Chavez\Documents\cqrrillo\compras.mdb" 
if  prod1=1 then
conn.execute "INSERT INTO pedidos(Codigopro,cantidad,precio) values('pla01'," & cint(cant) & ", 1800)"
end if
conn.close
set conn=nothing
response.redirect("index.html")

%>
</body>
</html>