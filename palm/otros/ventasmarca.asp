<%
on error resume next
dim rs, oConn, Sql
'------------ Coneccion 1 ------------------
Source ="SQLSERVER"
Catalog="todo"
Usuario="Flexline"
Password="corona"
set oConn = Server.CreateObject("ADODB.Connection")
oConn.open "Provider=SQLOLEDB;Data Source=" & Source & _
";Initial Catalog=" & Catalog & _
";User Id=" & Usuario & _
";Password=" & Password & _
";"

'Set oConn = server.createobject("ADODB.Connection")
'oConn.open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=todo;User Id=sa;Password=desakey;"
nuser = trim(request.querystring("nuser"))
nuser =right("000" & trim(nuser),3)

user = "aldo lopez"
SQL="SELECT nombre, NU_2 FROM Flexline.FX_vende_man WHERE (NU_2 = N'" & nuser & "')"
Set rs=oConn.execute(Sql)
user=rs.fields("nombre")


SQL="SELECT FX_PRODUCTO.TIPO AS Marca, SUM(Flexline.PY_Proyeccion.Meta) AS Meta, SUM(Flexline.PY_Proyeccion.Venta) AS Venta, " & _
"SUM(Flexline.PY_Proyeccion.Meta$) AS Meta$, SUM(Flexline.PY_Proyeccion.Venta$) AS Venta$ " & _
"FROM Flexline.PY_Proyeccion INNER JOIN " & _
"FX_PRODUCTO ON Flexline.PY_Proyeccion.Producto = FX_PRODUCTO.PRODUCTO " & _
"WHERE (Flexline.PY_Proyeccion.Vendedor_cliente = N'"& user &"') AND (Flexline.PY_Proyeccion.Periodo = CAST(YEAR(GETDATE()) AS nvarchar) "&_
"+ RIGHT('00' + CAST(MONTH(GETDATE()) AS nvarchar), 2)) AND " & _ 
"(FX_PRODUCTO.EMPRESA = 'desa') " & _
"GROUP BY FX_PRODUCTO.TIPO " & _
"ORDER BY FX_PRODUCTO.TIPO"
'response.write(sql)
set rs=oConn.execute(Sql)
call mireprot()
'----------------------------------------------------------------------------
sub mireprot()
%>
<html>
<head>
<title>Ventas Por Marca</title>
</head>

<body>

<div align="center"><strong>Ventas Por Marca </strong></div>

<font color="#CCCCCC" face="Arial" size="-6">
	<p align="center"><b>MES:</b><font style="text-transform:capitalize"><%= monthname(month(date)) %>
</font></p></font>

<table border="1" style="border-collapse:collapse" bordercolor="#000000">
<tr>
	<td alingh= "center"bgcolor="#000099"><CENTER><FONT SIZE="2" COLOR="#FFFFFF"><B>Marca</B></FONT></CENTER></td>
	<td alingh= "center"bgcolor="#000099"><CENTER><FONT SIZE="2" COLOR="#FFFFFF"><B>Meta</B></FONT></CENTER></td>
	<td alingh= "center"bgcolor="#000099"><CENTER><FONT SIZE="2" COLOR="#FFFFFF"><B>Venta</B></FONT></CENTER></td>
	<td alingh= "center"bgcolor="#000099"><CENTER><FONT SIZE="2" COLOR="#FFFFFF"><B>Meta$</B></FONT></CENTER></td>
	<td alingh= "center"bgcolor="#000099"><CENTER><FONT SIZE="2" COLOR="#FFFFFF"><B>Venta$</B></FONT></CENTER></td>
</tr>
<%
T1=0
T2=0
T3=0
T4=0

Do until rs.eof
T1=T1+cdbl(rs.fields("meta"))
T2=T2+cdbl(rs.fields("venta"))
T3=T3+cdbl(rs.fields("meta$"))
T4=T4+cdbl(rs.fields("venta$"))
%>
<tr>
	<td><%=replace(rs.fields("marca")," ","&nbsp;")%></td>
	<td><CENTER><%=formatnumber(rs.fields("meta"),0)%></CENTER></td>
	<td><CENTER><%=formatnumber(rs.fields("venta"),0)%></CENTER></td>
	<td><CENTER><%=formatnumber(rs.fields("meta$"),0)%></CENTER></td>
	<td><CENTER><%=formatnumber(rs.fields("venta$"),0)%></CENTER></td>
</tr>
<%
rs. movenext
loop
%>
<tr>
	<td>&nbsp;</td>
	<td><CENTER><B><%=formatnumber(T1,0)%></B></CENTER></td>
	<td><CENTER><B><%=formatnumber(T2,0)%></B></CENTER></td>
	<td><CENTER><B><%=formatnumber(T3,0)%></B></CENTER></td>
	<td><CENTER><B><%=formatnumber(T4,0)%></B></CENTER></td>
</tr>
</table>
<div align="center" class="Estilo2"><br>
  <br>
<input type="button" class="Estilo2" onClick="history.back()" value="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Volver&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"/>
</div>
<br>&nbsp;
<br>&nbsp;
<div align="center" class="Estilo4"><br>&nbsp;
Distribuci&oacute;n y Excelencia S.A.
</div>
</body>
</html>
<%
end sub 'mireprot()
'---------------------------------------------------------
%>