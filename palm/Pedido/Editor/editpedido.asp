<% Dim oConn, Vnd
set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=handheld;User Id=sa;Password=desakey"
nuser = Request.Querystring("nuser")
SQL = "SELECT nombre AS vendedor, NU_2 AS Nuser "&_
	  "FROM Flexline.FX_vende_man "&_
	  "WHERE (NU_2 = '" & nuser & "')"
Set rs0 = oConn.Execute(SQL)

Vnd = rs0.fields("vendedor")
%>
<html>
<head>
<title>Editar Pedido</title>
</head>
<body>
<TABLE border='1'>
	<TR bgcolor='#9999DB'>
		<TD align='center'><B>Edit</B></TD>
		<TD align='center'><B>Nota</B></TD>
		<TD align='center'><B>Pedido</B></TD>
		<TD align='center'><B>Cliente</B></TD>
		<TD align='center'><B>Estado</B></TD>
	</TR>
<%
'nuser=29
sql="select * FROM sqlserver.desaerp.dbo.PED_PEDIDOSENC WHERE (idvendedor = " & cint(nuser) & ") AND (sw_estado in('I','P'))"
Set rs = oConn.Execute(sql)
x=1
do until rs.eof
cliente=rs.fields("idcliente") & " " & rs.fields("idsucursal")
if (x mod 2) = 0 then
'if x=1 then
	colorfondo="#C0C0C0"
else
	colorfondo="#E6E6E6"
end if
%>
	<TR bgcolor='<%=colorfondo%>'>
		<TD><A HREF="elimina.asp?nuser=<%=nuser%>&np=<%=rs.fields("Numero_pedido")%>">eliminar</A></TD>
		<TD><%=right(rs.fields("Numero_pedido"),4)%></TD>
		<TD><%=rs.fields("Numero_pedido")%></TD>
		<TD><%=cliente %></TD>
		<TD><%=rs("sw_estado") %></TD>
	</TR>
<%
rs.movenext
x=x+1
loop
'-------
function iif(bolen,vtrue,vfalse)
	if bolean then
		iif = vtrue
	else
		iif =vfalse
	end if
end function
%>
</TABLE>
</body>
