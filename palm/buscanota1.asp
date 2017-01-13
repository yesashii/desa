<%
'on error resume next
'.:: Conexion ::.
set oConn=server.Createobject("ADODB.Connection")
oConn.open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=handheld;User Id=sa;Password=desakey"
Dim user 'Nombre Vendedor
Dim nuser'Numero Vendedor 
Dim nota 'Numero Notaventa
Dim np   'Numero Notaventa
Dim op   'año "AA"
Dim sql_cliente
Dim sql_LineasProducto
Dim sql_notapedido
dim direccion
'Dim


if request.Querystring="" then
	call SinDatos()
else
	call buscadatos()
	call NotaPedido()
end if
'---------------------------------------------------------------------------------------------
sub buscadatos()
	'.:: Variables ::.
	user = request.Querystring("user")
	nota = request.Querystring("nota")
	np   = request.Querystring("np"  )
	if len(np)>1 then
		nota  = right(np,4)
		nuser = cint(left(right(np,7),3))
		sql="select nombre from SQLSERVER.desaerp.dbo.DIM_VENDEDORES where idvendedor=" & nuser
		set rs=oConn.execute(Sql) 
		user  = rs.fields(0)
		ao=left(np,2)
	end if
	if len(ao)=0 then ao=right(cstr(year(date-5)),2)

	sql_cliente="SELECT c.RazonSocial, c.CodLegal, CAST(SUBSTRING(c.CtaCte, CHARINDEX('-', c.CtaCte) + 3, LEN(c.CtaCte) - CHARINDEX('-', c.CtaCte) + 3) AS numeric) AS local,  c.ListaPrecio, p.* " & _
	"FROM handheld.Flexline.FX_PEDIDO_PDA as P INNER JOIN Flexline.CtaCte  as C ON p.Cliente = c.CtaCte " & _
	"WHERE (p.nota =  "& nota &") AND (p.vendedor = N'"& user &"') AND (c.Empresa = 'DESA') AND (c.TipoCtaCte = 'CLIENTE') and (left(p.fecha,2) = '"& ao &"')"

	 sql_LineasProducto="SELECT L.*, P.Glosa " & _
	"FROM  sqlserver.desaerp.dbo.PED_PEDIDOSDET as L " & _
	" inner join serverdesa.BDFlexline.flexline.producto as P on P.producto=L.Producto " & _
	"WHERE     (L.numero_pedido = N'" & np & "') " & _
	"and p.empresa='desa' " & _
	"ORDER BY L.linea_pedido"

	sql_notapedido="SELECT N.* " & _
	"FROM  sqlserver.desaerp.dbo.PED_PEDIDOSENC as N " & _
	"WHERE     (N.numero_pedido = N'" & np & "') "
	
'set rs=oConn.execute(sql)
	

	set rs=oConn.execute(Sql)
	if rs.eof then 
		sql=replace(sql,"handheld.Flexline.FX_PEDIDO_PDA","handheld.dbo.DesaERP_PEDIDOS")
		set rs=oConn.execute(Sql)
	end if
	'rsfecha=rs.fields("fecha")
End sub 'buscadatos()
'---------------------------------------------------------------------------------------------
Sub NotaPedido()
	%><html>
	<head>
	<title>Nota de Pedido</title>
	<style type="text/css">
	<!--
	body {
		background-color: #FFFFFF;
		font-family: Arial;
		/*font-size: 10px;*/
		font-size: 8;
	}
	td {
		font-size: 12px;
		/*border: 0;
		border-width:0;*/
		border-collapse: collapse;
		/*text-align:center;*/
	}
	th {
		background-color: #000033;
		font-family: Arial;
		font-size: 12;
		color: #FFFFFF;

	}
	table {
		border: 0;
		border-width:0;
		border-collapse: collapse;
	}
	-->
	</style>
	</head>

	<body>
	<CENTER>
	<TABLE width="650" name="marco"><TR><TD>
<!-- Datos Cliente -->
<% set rs=oConn.execute(sql_cliente) %>
	<TABLE>
	<TR>
		<TD width="100"><B>Codigo&nbsp;Legal</B></TD>
		<TD>:&nbsp;<%=rs("codlegal")%>&nbsp;&nbsp;&nbsp;local&nbsp;:&nbsp;<B><%=rs("local")%></B></TD>
	</TR>
	<TR>
		<TD><B>Nombre</B></TD>
		<TD>:&nbsp;<%=pc(rs("RazonSocial"))%></TD>
	</TR>
	<%	Sql="SELECT * FROM serverdesa.BDFlexline.flexline.CtaCteDirecciones WHERE (CtaCte = '"& rs.fields("cliente") &"') AND "&_
	"(Principal <> 's') AND (TipoCtaCte = 'cliente') AND (Empresa = 'desa')"
	set rsd=oConn.execute(Sql)
	if not rsd.eof then direccion=replace(rsd("direccion"),"|","-") & " (" & rsd("comuna") & ")" %>
	<TR>
		<TD><B>Direccion</B></TD>
		<TD>:&nbsp;<%=pc(direccion)%></TD>
	</TR>
	<TR>
		<TD><B>Vendedor</B></TD>
		<TD>:&nbsp;<%=pc(user)%></TD>
	</TR>
	<TR>
		<TD><B></B></TD>
		<TD></TD>
	</TR>
	</TABLE>
<!-- Datos Documento --><HR>
<% set rs=oConn.execute(sql_notapedido) %>
	<TABLE>
	<TR>
		<TD width="100"><B>Nota&nbsp;Pedido</B></TD>
		<TD>:&nbsp;<%=rs("Numero_pedido")%></TD>
	</TR>
	<TR>
		<TD width="100"><B>Estado</B></TD>
		<TD>:&nbsp;
		<% estado=ucase(rs("sw_estado"))
		if estado="I" then estado="Ingresado"
		if estado="P" then estado="En Proceso"
		if estado="F" then estado="Facturado"
		if estado="R" then estado="Rechazado"
		if estado="N" then estado="Nulo"
		response.write(estado)
		%>
		</TD>
	</TR>
	<TR>
		<TD><B>Orden&nbsp;Compra</B></TD>
		<TD>:&nbsp;<%=rs("Orden_compra")%></TD>
	</TR>
	<TR>
		<TD><B>Correlativo Flex</B></TD>
		<TD>:&nbsp;<%=rs("Correlativo_flex")%></TD>
	</TR>
	<TR>
		<TD><B>Fecha&nbsp;Solicitud</B></TD>
		<TD>:&nbsp;<%=ffecha(rs("fecha_pedido")) & "&nbsp;&nbsp;&nbsp;Hora&nbsp;:&nbsp;" & fhora(rs("hora_pedido")) %></TD>
	</TR>
		<TR>
		<TD><B>Entregar&nbsp;Fecha</B></TD>
		<TD>:&nbsp;<%=ffecha(rs("fecha_Facturacion")) %></TD>
	</TR>
	<TR>
		<TD><B>Vendedor</B></TD>
		<TD>:&nbsp;<%=pc(user)%></TD>
	</TR>
	<TR>
		<TD><B></B></TD>
		<TD></TD>
	</TR>
	</TABLE>
<!-- Lineas Producto -->
<TABLE width="100%" border="1" bordercolor="#808080" cellpadding="0" cellspacing="0" style="border-color: #000000;border-style: solid">
<TR>
	<TD style="border: 0"></TD>
	<TD style="border: 0"></TD>
	<TH colspan="2" style="border-width: 1px;border-right-width: 2px;">Cantidad</TH>
	<TD style="border: 0"></TD>
	<TD style="border: 0"></TD>
</TR>
<TR>
	<TH>Codigo</TH>
	<TH>Descripcion</TH>
	<TH>Pedida</TH>
	<TH>Desp.</TH>
	<TH>Precio</TH>
	<TH>Desc</TH>
</TR>
<% 
set rs=oConn.execute(sql_LineasProducto)
do until rs.eof %>
<TR border="1">
	<TD align="center"><%=rs("Producto"               )%></TD>
	<TD align="left"  >&nbsp;<%=pc(rs("glosa"        ))%></TD>
	<TD align="center"><%=pc(rs("Cantidad_Pedida"    ))%></TD>
	<TD align="center"><%=pc(rs("Cantidad_Despachada"))%></TD>
	<TD align="center"><%=pc(rs("Precio_unitario"    ))%></TD>
	<TD align="center">-<%=rs("Descuento_Unitario"    )%></TD>
</TR>
<% rs.movenext
loop%>
</TABLE>
<!-- Etado Pedido --><BR>
<%
ap_fin="Sin Problemas"
ap_pre="Sin Problemas"
ap_sto="Sin Problemas"

sql="SELECT * " & _
	"FROM  sqlserver.desaerp.dbo.PED_PEDIDOSENC as N " & _
	"WHERE     (N.numero_pedido = N'" & np & "') " & _
	"and vb_aprfin='N'"
set rs=oConn.execute(sql)
if not rs.eof then ap_fin="Problemas Financieros"

sql="SELECT * " & _
	"FROM  sqlserver.desaerp.dbo.PED_PEDIDOSENC as N " & _
	"inner join sqlserver.desaerp.dbo.PED_PEDIDOSDET as L on N.Numero_pedido=L.Numero_pedido " & _
	"WHERE     (N.numero_pedido = N'" & np & "') " & _
	"and N.vb_aprfin='S' and N.sw_estado in('F','R') and L.vb_precio='N' " & _
	"order by L.linea_pedido"
set rs=oConn.execute(sql)
if not rs.eof then 'ap_fin="Problemas Precio"
	ap_pre="<TABLE>"
	do until rs.eof
		ap_pre=ap_pre & "<TR><TD>" & rs("producto") & "</TD><TD>Descuento No Aprobado</TD></TR>"
	rs.movenext
	loop
	ap_pre=ap_pre & "</TABLE>"
End if

sql="SELECT * " & _
	"FROM  sqlserver.desaerp.dbo.PED_PEDIDOSENC as N " & _
	"inner join sqlserver.desaerp.dbo.PED_PEDIDOSDET as L on N.Numero_pedido=L.Numero_pedido " & _
	"WHERE     (N.numero_pedido = N'" & np & "') " & _
	"and N.vb_aprfin='S' and N.sw_estado in('F','R') and L.vb_precio='N' " & _
	"order by L.linea_pedido"
set rs=oConn.execute(sql)
if not rs.eof then 'ap_fin="Problemas Precio"
	ap_sto="<TABLE>"
	do until rs.eof
		ap_sto=ap_sto & "<TR><TD>" & rs("producto") & "</TD><TD>Sin Stock suficiente</TD></TR>"
	rs.movenext
	loop
	ap_sto=ap_sto & "</TABLE>"
End if



%>
<TABLE>
<TR>
	<TD width="20"></TD>
	<TD width="20"></TD>
	<TD></TD>
</TR>
<TR>
	<TD colspan="3"><B><U>Posibles Causas</U></B></TD>
</TR>
<TR>
	<TD></TD>
	<TD><B>De tipo Financiero</B></TD>
	<TD></TD>
</TR>
<TR>
	<TD></TD>
	<TD>&nbsp;-<%=ap_fin%></TD>
	<TD></TD>
</TR>
<TR>
	<TD></TD>
	<TD><B>De tipo Precio</B></TD>
	<TD></TD>
</TR>
<TR>
	<TD></TD>
	<TD>&nbsp;-<%=ap_pre%></TD>
	<TD></TD>
</TR>
<TR>
	<TD></TD>
	<TD><B>De tipo Disponibilidad</B></TD>
	<TD></TD>
</TR>
<TR>
	<TD></TD>
	<TD>&nbsp;-<%=ap_sto%></TD>
	<TD></TD>
</TR>
</TABLE>
	</TD></TR></TABLE>
	</CENTER>
	</body>
	<%
End Sub 'NotaPedido()
'---------------------------------------------------------------------------------------------
function pc(texto)
sString=texto
	Dim sWhiteSpace, bCap, iCharPos, sChar 
  sWhiteSpace = Chr(32) & Chr(9) & Chr(13) 
  sString = LCase(sString)
  bCap = True 
  For iCharPos = 1 to Len(sString) 
    sChar = Mid(sString, iCharPos, 1) 
    If bCap = True Then 
      sChar = UCase(sChar) 
    End If 
    ProperCase = ProperCase + sChar 
    If InStr(sWhiteSpace, sChar) Then 
      bCap = True 
    Else 
      bCap = False 
    End If 
  Next  
texto=ProperCase
	if isnumeric(trim(texto)) then texto=formatnumber(cdbl(texto),0)
	pc=texto
end function 'pc
'---------------------------------------------------------------------------------------------
Sub SinDatos()
	%><html>
	<head>
	<title>DESA</title>
	</head><body>
	<p align='center'>
	<BR><BR>Se necesitan datos
	<BR><input type="button" value="<< Atras" onclick="history.back()">
	</p>
	</body>
	</html><%
End Sub 'SinDatos()
'---------------------------------------------------------------------------------------------
function ffecha(fecha)
	dd=right(fecha,2)
	mm=left(right(fecha,4),2)
	yyyy=left(fecha,4)
	ffecha=dd & "/" & mm & "/" & yyyy
End function 'ffecha
'---------------------------------------------------------------------------------------------
function fhora(hora)
	hora=right("000000" & hora,6)
	hh=left(hora,2)
	mm=left(right(hora,4),2)
	ss=right(hora,2)
	fhora=hh & ":" & mm '& " :" & ss
End function 'fhora
'---------------------------------------------------------------------------------------------
%>