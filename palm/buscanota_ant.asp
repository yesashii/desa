<%
on error resume next
set oConn=server.Createobject("ADODB.Connection")
set oConn1=server.Createobject("ADODB.Connection")
oConn.open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=handheld;User Id=sa;Password=desakey"
oConn1.open "Provider=SQLOLEDB;Data Source=serverdesa;Initial Catalog=BDFlexline;User Id=sa;Password=desakey"'serverdesa
if request.Querystring = "" then
response.write("<html>"& chr(13) &"<head>"& chr(13) &"<title>DESA</title>"& chr(13) &"</head>")
response.write(chr(13) &"<body>"& chr(13) &"<p align='center'>se necesitan datos</p>")
response.write(chr(13) &"</body>"& chr(13) &"</html>")
else

user = request.Querystring("user")
nota = request.Querystring("nota")
np = request.Querystring("np")
if len(np)>1 then
	nota  = right(np,4)
	nuser = cint(left(right(np,7),3))
	sql="select nombre from SQLSERVER.desaerp.dbo.DIM_VENDEDORES where idvendedor=" & nuser
	set rs=oConn.execute(Sql) 
	user  = rs.fields(0)
	ao=left(np,2)
	'response.write(user & "<br>")
	'response.write(nota & "<br>")
end if
if len(ao)=0 then ao="08"

'Sql = "SELECT Flexline.CtaCte.RazonSocial, Flexline.CtaCte.CodLegal, CAST(SUBSTRING(Flexline.CtaCte.CtaCte, CHARINDEX('-', Flexline.CtaCte.CtaCte) + 3, LEN(Flexline.CtaCte.CtaCte) - CHARINDEX('-', Flexline.CtaCte.CtaCte) + 3) AS numeric) AS local,  Flexline.CtaCte.ListaPrecio, handheld.Flexline.FX_PEDIDO_PDA.* "&_
'"FROM handheld.Flexline.FX_PEDIDO_PDA INNER JOIN Flexline.CtaCte ON handheld.Flexline.FX_PEDIDO_PDA.Cliente = Flexline.CtaCte.CtaCte "&_
'"WHERE (handheld.Flexline.FX_PEDIDO_PDA.nota = "& nota &") AND (handheld.Flexline.FX_PEDIDO_PDA.vendedor = N'"& user &"') AND (Flexline.CtaCte.Empresa = 'DESA') AND (Flexline.CtaCte.TipoCtaCte = 'CLIENTE') and (left(handheld.Flexline.FX_PEDIDO_PDA.fecha,2) = '08')"

sql="SELECT c.RazonSocial, c.CodLegal, CAST(SUBSTRING(c.CtaCte, CHARINDEX('-', c.CtaCte) + 3, LEN(c.CtaCte) - CHARINDEX('-', c.CtaCte) + 3) AS numeric) AS local,  c.ListaPrecio, p.* " & _
"FROM handheld.Flexline.FX_PEDIDO_PDA as P INNER JOIN Flexline.CtaCte  as C ON p.Cliente = c.CtaCte " & _
"WHERE (p.nota =  "& nota &") AND (p.vendedor = N'"& user &"') AND (c.Empresa = 'DESA') AND (c.TipoCtaCte = 'CLIENTE') and (left(p.fecha,2) = '"& ao &"')"

'response.write(sql)
set rs=oConn.execute(Sql)
if rs.eof then 
	sql=replace(sql,"handheld.Flexline.FX_PEDIDO_PDA","handheld.dbo.DesaERP_PEDIDOS")
	set rs=oConn.execute(Sql)
end if
rsfecha=rs.fields("fecha")
%>
<html>
<head>
<title>B&uacute;squeda por n&ordm; nota de venta</title>

<style type="text/css">
<!--
body {
	background-color: #E5E5E5;
	font-family: Arial;
	font-size: 12px;
}
-->
</style>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head>

<body>
<table cellpadding="0" cellspacing="0">
	<tr onclick="window.open('detallera.asp?nota=<%=np%>')">
		<td colspan="2" nowrap><strong>Nota Venta</strong></td>
		<td width="82%" nowrap><strong>:</strong>
			<%= rs.fields("nota") & " : " & np %></td>
	</tr>
	<tr>
		<td colspan="2" nowrap><strong>Orden de Compra N&deg;
</strong></td>
		<td nowrap><strong>:</strong> 
			<%= rs.fields("oc") %></td>
	</tr>
	<tr>
		<td colspan="2" nowrap><strong>Fecha Solicitud </strong></td>
		<td nowrap><strong>:</strong> 
<% mifecha = right(rs.fields("fecha"),2) &"/"& mid(rs.fields("fecha"),3,2) & "/20" & left(rs.fields("fecha"),2)
response.write(mifecha)
%></td>
	</tr>
	<tr>
		<td colspan="2" nowrap><strong>Hora Solicitud </strong></td>
		<td nowrap><strong>:</strong> 
<% mihora = left(rs.fields("hora"),2) &":"& mid(rs.fields("hora"),3,2) & ":" & right(rs.fields("hora"),2)
response.write(mihora)
%>
</td>
	</tr>
<% if not (len(rs.fields("fechaentrega")) = 0 or len(rs.fields("fechaentrega"))> 8)then %>	
	<tr>
		<td colspan="2" nowrap><strong>Fecha solic. Despacho</strong></td>
		<td nowrap><strong>:</strong> 
<% mifecha2 = right(rs.fields("fechaentrega"),2) &"/"& mid(rs.fields("fechaentrega"),5,2) & "/" & left(rs.fields("fechaentrega"),4)
if cdate(mifecha) > cdate("23/10/2006") then
response.write(cdate(mifecha2))
else
response.Write(mifecha2)
end if
%></td>
	</tr>
<% end if %>
	<tr>
		<td colspan="3" nowrap><hr color="#666666"></td>
	</tr>
	<tr>
		<td colspan="2" nowrap><strong>Cliente</strong></td>
		<td nowrap><strong>:</strong> 
			<%= rs.fields("codlegal") %></td>
	</tr>
	<tr>
		<td colspan="2" nowrap><strong>Nombre</strong></td>
		<td nowrap><strong>:</strong> 
			<%= rs.fields("razonsocial") %></td>
	</tr>
	<tr>
		<td colspan="2" nowrap><strong>Local</strong></td>
		<td nowrap><strong>:</strong>  
			<%= rs.fields("local") %></td>
	</tr>
	<tr>
		<td colspan="2" nowrap><strong>Direcci&oacute;n</strong></td>
		<% Sql="SELECT * FROM CtaCteDirecciones WHERE (CtaCte = '"& rs.fields("cliente") &"') AND "&_
"(Principal <> 's') AND (TipoCtaCte = 'cliente') AND (Empresa = 'desa')"
set rs1=oConn1.execute(Sql)
%>
		<td nowrap><strong>:</strong> 
			<%= rs1.fields("direccion") %></td>
	</tr>
	<tr>
		<td colspan="2" nowrap><strong>Vendedor</strong></td>
		<td nowrap><strong>: 
			</strong>
			<%= rs.fields("vendedor") %></td>
	</tr>
	<tr>
		<td colspan="3">&nbsp;</td>
	</tr>
</table>
<hr color="#666666">
<table align="center" cellpadding="0" cellspacing="0" id="Main">
	<tr>
		<td align="center" nowrap><strong>
			&nbsp;
			C&oacute;digo
			&nbsp;
		</strong></td>
		<td nowrap><strong>
			&nbsp;
		Descripci&oacute;n
		&nbsp;
		</strong></td>
		<td align="center" nowrap><strong>
			&nbsp;
		Cantidad
		&nbsp;</strong></td>
		<td align="center" nowrap><strong>
			&nbsp;
		Precio
		&nbsp;
		</strong></td>
		<td align="center" nowrap><strong>
			&nbsp;
		Descuento
		&nbsp;
		</strong></td>
	</tr>
<% for x = 1 to 12 
if not rs.fields("producto" & right("00" & x,2)) = "" then
if x mod 2 = 0 then 
color = "#FFFFCC"
else
color = "#CCFFCC"
end if
Sql2 ="SELECT PRODUCTO.PRODUCTO, PRODUCTO.GLOSA, PRODUCTO.FAMILIA, ListaPrecioD.Valor "&_
"FROM serverdesa.BDFlexline.flexline.ListaPrecio ListaPrecio INNER JOIN serverdesa.BDFlexline.flexline.ListaPrecioD ListaPrecioD ON ListaPrecio.Empresa = ListaPrecioD.Empresa AND ListaPrecio.IdLisPrecio = ListaPrecioD.IdLisPrecio INNER JOIN serverdesa.BDFlexline.flexline.PRODUCTO PRODUCTO ON ListaPrecioD.Empresa = PRODUCTO.EMPRESA AND ListaPrecioD.Producto = PRODUCTO.PRODUCTO "&_
"WHERE (PRODUCTO.EMPRESA = 'DESA') AND (PRODUCTO.TIPOPRODUCTO = 'EXISTENCIA') AND (ListaPrecio.LisPrecio = '"& rs.fields("listaprecio") &"') AND (PRODUCTO.PRODUCTO = '"& rs.fields("producto" & right("00" & x,2)) &"')"

set rs2=oConn.execute(Sql2)
%> 
	<tr bgcolor="<%= color %>">
		<td align="center" nowrap><strong>
			<%= rs.fields("producto" & right("00" & x,2)) %>
		</strong></td>
		<td nowrap><strong>
			&nbsp;
			<%= replace(left(rs2.fields("glosa"),40)," ","&nbsp;")%>
		</strong></td>
		<td align="center" nowrap><strong>
			<%= rs.fields("cantidad" & right("00" & x,2)) %>
		</strong></td>
		<td align="center" nowrap><strong>
			<%= formatnumber(rs2.fields("valor"),0) %>
		</strong></td>
		<td align="center" nowrap><strong>
			<%= rs.fields("descuento" & right("00" & x,2)) %>
		%</strong></td>
	</tr>
<%
else
'exit for
end if
next %>
</table>
<hr color="#666666">
<% if not len(rs.fields("OBS"))<= 1 then %>
<table cellpadding="0" cellspacing="0">
<tr>
	<td nowrap>&nbsp;
		<strong>Observaciones</strong>: </td>
</tr>
<tr>
	<td nowrap style="text-transform:uppercase">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		&nbsp;&nbsp;&nbsp;&nbsp;<%= rs.fields("OBS") %>&nbsp;</td>
</tr>
</table>
<%end if
'call DatosFactuRA()
%>
 
</body>
</html>
<%
end if
'------------------------------------------------------------------------------------------------------
Sub DatosFactuRA()
on error resume next
sql="select numero from handheld.flexline.FX_vende_man where nombre='" & user & "'"
Set rs=oConn.execute(Sql)
nuser=rs.fields(0)
npedido= left(rsfecha,4) & right("000" & nuser,3) & right("0000" & nota,4)
sql="select " & _
"	usuario_aprfin as Usuario, " & _
"	vb_aprfin as  'vb Financiero', " & _
"	vb_cremax as 'Credito Maximo',  " & _
"	vb_deucta as 'Deuda Activa',  " & _
"	vb_provig as 'Protestos Vigentes', " & _
"	vb_prohis as 'Protestos Historicos',  " & _
"	vb_diaatr as 'Dias de Atraso',  " & _
"	vb_precio as 'VB Precio',  " & _
"	vb_stock  as 'VB Stock',  " & _
"	replace(sw_estado,'F','Facturado') as 'Estado Pedido' " & _
"from sqlserver.desaerp.dbo.PED_PEDIDOSENC " & _
"where numero_pedido='" & npedido & "'"
Set rs=oConn.execute(Sql)
do until rs.eof
	response.write("<HR><TABLE>")
	for x=0 to rs.fields.count
		response.write("<TR><TD>" & rs.fields(x).name & "</TD>")
		response.write("<TD>" & rs.fields(x).value & "</TD></TR>")
	next
	response.write("</TABLE>")
rs.movenext
loop
End Sub 'DatosFactuRA()
'------------------------------------------------------------------------------------------------------

%>