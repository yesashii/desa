<%
dim rs, oConn, Sql
'------------coneccion 1------------------
Set oConn = server.createobject("ADODB.Connection")

Set oConn1 = server.createobject("ADODB.Connection")

oConn1.open "Provider=SQLOLEDB;Data Source=serverdesa;Initial Catalog=BDFlexline;User Id=sa;Password=desakey;"

oConn.open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=flexline2;User Id=sa;Password=desakey;"

nuser = trim(request.querystring("nuser"))

nuser =right(nuser,2)

SQL1="select * from BDFlexline.dbo.PALM_VENDEDOR PALM_VENDEDOR where NV='" & nuser & "'"
Set rs1=oConn1.execute(Sql1)
user=rs1.fields("descripcion")

'response.write(user)

Sql="SELECT dbo.FX_Rent_Acum_2006.Marca, SUM(dbo.FX_Rent_Acum_2006.Venta) AS Venta, dbo.FX_Rent_Acum_2006.Vendedor_factura, "&_
"dbo.FX_Rent_Acum_2006.PeriodoLibro, SUM(dbo.FX_Rent_Acum_2006.Cantidad / dbo.PALM_PRODUCTO.FACTORALT) AS Cajas_Vta, "&_
"SUM(Todo.Flexline.FX_Forecast_Meta.Cajas) AS Cajas, SUM(Todo.Flexline.FX_Forecast_Meta.Cajas * Todo.Flexline.FP_SKU.ValorCaja) "&_
"AS META_vta "&_
"FROM dbo.PALM_PRODUCTO INNER JOIN "&_
"dbo.FX_Rent_Acum_2006 ON dbo.PALM_PRODUCTO.PRODUCTO = dbo.FX_Rent_Acum_2006.Producto INNER JOIN "&_
"Todo.Flexline.FX_Forecast_Meta ON dbo.FX_Rent_Acum_2006.Vendedor_factura = Todo.Flexline.FX_Forecast_Meta.Vendedor AND "&_
"dbo.FX_Rent_Acum_2006.Producto = Todo.Flexline.FX_Forecast_Meta.Producto AND "&_
"dbo.FX_Rent_Acum_2006.PeriodoLibro = Todo.Flexline.FX_Forecast_Meta.Periodo INNER JOIN "&_
"Todo.Flexline.FP_SKU ON Todo.Flexline.FX_Forecast_Meta.Producto = Todo.Flexline.FP_SKU.Producto "&_
"GROUP BY dbo.FX_Rent_Acum_2006.Marca, dbo.FX_Rent_Acum_2006.Vendedor_factura, dbo.FX_Rent_Acum_2006.PeriodoLibro "&_
"HAVING (dbo.FX_Rent_Acum_2006.Vendedor_factura = N'"& user &"') AND (dbo.FX_Rent_Acum_2006.PeriodoLibro = CAST(YEAR(GETDATE()) AS nvarchar) "&_
"+ RIGHT('00' + CAST(MONTH(GETDATE()) AS nvarchar), 2)) "&_
"ORDER BY dbo.FX_Rent_Acum_2006.Marca"

SQL="SELECT     Todo.Flexline.PY_Proyeccion.Periodo, dbo.PRODUCTO.TIPO AS Marca, SUM(Todo.Flexline.PY_Proyeccion.Meta) AS Meta, " & _
"SUM(Todo.Flexline.PY_Proyeccion.Venta) AS Venta, SUM(Todo.Flexline.PY_Proyeccion.Meta$) AS Meta$, SUM(Todo.Flexline.PY_Proyeccion.Venta$) " & _
"AS Venta$ " & _
"FROM         Todo.Flexline.PY_Proyeccion INNER JOIN " & _
"dbo.PRODUCTO ON Todo.Flexline.PY_Proyeccion.Producto = dbo.PRODUCTO.PRODUCTO " & _
"GROUP BY Todo.Flexline.PY_Proyeccion.Periodo, Todo.Flexline.PY_Proyeccion.Vendedor_Factura, dbo.PRODUCTO.TIPO, dbo.PRODUCTO.EMPRESA " & _
"HAVING      (dbo.PRODUCTO.EMPRESA = 'desa') AND (Todo.Flexline.PY_Proyeccion.Periodo = CAST(YEAR(GETDATE()) AS nvarchar) "&_
"+ RIGHT('00' + CAST(MONTH(GETDATE()) AS nvarchar), 2)) AND " & _
"(Todo.Flexline.PY_Proyeccion.Vendedor_Factura = N'"& user &"')"
''antigua
'Sql	="SELECT Marca, SUM(Venta) AS venta, Vendedor_factura, PeriodoLibro, SUM(Cantidad) AS botellas FROM dbo.FX_Rent_Acum_2006 GROUP BY Marca, Vendedor_factura, PeriodoLibro HAVING (Vendedor_factura = N'" & user & "') AND (PeriodoLibro = CAST(YEAR(GETDATE()) AS nvarchar) + RIGHT('00' + CAST(MONTH(getdate()) AS nvarchar), 2)) ORDER BY Marca"
'response.write(sql)

set rs=oConn.execute(Sql)
%>
<html>
<head>
<title>Ventas Por Marca</title>
<style type="text/css">
<!--
.Estilo1 {
	color: #CCCCCC;
	font-weight: bold;
	font-size: 12px;
	font-family: arial;
}
.Estilo2 {
	font-family: Arial;
	font-size: 12px;
	color: #000000;
}
.Estilo4 {color: #CCCCCC; font-weight: bold; font-size: 9px; font-family: arial; }
-->
</style>
</head>

<body>
<div align="center"><strong>Ventas Por Marca </strong></div>
<font color="#CCCCCC" face="Arial" size="-6"><p align="center"><b>MES:</b><font style="text-transform:capitalize"><%= monthname(month(date)) %></font></p></font>
<table width="100%" border="1" bordercolor="#000000" style="border-collapse:collapse">
<tr>
    <td align="center" nowrap bgcolor="#000099" class="Estilo1">Marca</td>
    <td align="center" nowrap bgcolor="#000099" class="Estilo1">Meta</td>
    <td align="center" nowrap bgcolor="#000099" class="Estilo1">Venta</td>
    <td align="center" nowrap bgcolor="#000099" class="Estilo1">Meta$</td>
  	<td align="center" nowrap bgcolor="#000099" class="Estilo1">Venta$</td>
</tr>
<%
bot=0
vta=0
met=0
do until rs.eof
%>
  <tr>
    <td nowrap class="Estilo2"><%=rs.fields("marca")%></td>
    <td align="center" nowrap class="Estilo2"><%=formatnumber(cdbl(rs.fields("meta")),0) %></td>
    <td align="center" nowrap class="Estilo2"><%= formatnumber(cdbl(rs.fields("venta")),0) %>&nbsp;</td>
    <td class="Estilo2" align="right"><%=formatnumber(cdbl(rs.fields("meta$")),0) %></td>
  	<td class="Estilo2" align="right"><%=formatnumber(cdbl(rs.fields("Venta$")),0) %>&nbsp;</td>
  </tr>
<%
bot= bot + cdbl(rs.fields("meta"))
met= met + cdbl(rs.fields("venta"))
vta= vta + cdbl(rs.fields("meta$"))
tmet = tmet + cdbl(rs.fields("venta$"))
rs.movenext
loop
%>
  <tr>
    <td nowrap class="Estilo2" align="right"><b>Total:&nbsp;</b></td>
    <td align="center" nowrap class="Estilo2"><B><%= formatnumber(bot,0) %></B></td>
    <td align="center" nowrap class="Estilo2"><B><%= formatnumber(met,0) %>&nbsp;</B></td>
    <td class="Estilo2" align="right"><B><%= formatnumber(vta,0) %></B></td>
  	<td class="Estilo2" align="right"><B><%= formatnumber(tmet,0) %></B></td>
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