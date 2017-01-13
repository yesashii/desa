<!-- <CENTER><FONT SIZE="2" face="vrdana" COLOR=""><B>76.055.635-1 - Distribuidora La Cav LTDA</B></FONT></CENTER>
<HR> -->
<%'on error resume next
'.:: Conexion ::.
set oConn=server.Createobject("ADODB.Connection")
oConn.open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=handheld;User Id=sa;Password=desakey"
Dim user 'Nombre Vendedor
Dim nuser'Numero Vendedor 
Dim nota 'Numero Notaventa
Dim np   'Numero Notaventa
Dim op   'año "AA"
Dim sqlduplica
Dim sql_cliente
Dim sql_LineasProducto
Dim sql_notapedido
Dim direccion
dim idempresa
Dim duplica
'dim empresa

if request.Querystring="" then
	call SinDatos()
else
	call buscadatos()
	
end if
'---------------------------------------------------------------------------------------------
sub validaempresa()
	if len(idempresa)=0 then
		
	end if
end sub
'---------------------------------------------------------------------------------------------
sub buscadatos()
	'on error resume next
	'.:: Variables ::.
	empresa="desa"
	user = request.Querystring("user")
	nota = request.Querystring("nota")
	np   = request.Querystring("np"  )
	idempresa=request.Querystring("idempresa")
	duplica=request.Querystring("duplica")

	if len(duplica)<1 then duplica=0
	
'---------DUPLICA EL PEDIDO--------------
if duplica=1 then
	
	sqlduplica="Exec dbo.DuplicarPedido '" & np & "', " & idempresa
	
	oConn.execute(sqlduplica)
	
	'response.write (sqlduplica)
	
end if
'---------DUPLICA EL PEDIDO--------------

'valida empresa
	if len(idempresa)=0  and len(np)>0 then
	
		sql="select * FROM sqlserver.desaerp.dbo.PED_PEDIDOSENC where numero_pedido='" & np & "'"
		set rs=oConn.execute(Sql) 
		xn=0
		xempresa=""
		do until rs.eof
			xempresa=xempresa & "|"
			idempresa=rs("idempresa")
		rs.movenext
		xn=xn+1
		loop

		'1: es ok
		if xn=0 then
			%><CENTER>NO SE ENCONTRO PEDIDO</CENTER><%
			exit sub
		end if
		if xn>1 then
		%><CENTER><BR><BR><BR>Elija empresa<BR>
			<FORM METHOD=POST ACTION="buscanota.asp?idempresa=1&np=<%=np%>">
				<INPUT TYPE="submit" value="DESA">
			</FORM>
			<FORM METHOD=POST ACTION="buscanota.asp?idempresa=4&np=<%=np%>">
				<INPUT TYPE="submit" value="LACAV">
			</FORM>
			<FORM METHOD=POST ACTION="buscanota.asp?idempresa=3&np=<%=np%>">
				<INPUT TYPE="submit" value="DESAZOFRI">
			</FORM>
		</CENTER><%
			exit sub
		end if
	end if
	
	
	if idempresa=1 then empresa="DESA"
	if idempresa=4 then empresa="LACAV"
	if idempresa=3 then empresa="DESAZOFRI"


	if len(np)>1 then
		nota  = right(np,4)
		nuser = cint(left(right(np,7),3))
		sql="select nombre from SQLSERVER.desaerp.dbo.DIM_VENDEDORES where idvendedor=" & nuser & " and idempresa=" & idempresa 
		set rs=oConn.execute(Sql) 
		user  = rs.fields(0)
		ao=left(np,2)
	else
		if len(nota)>0 and len(user)>0 then
		sql="select idvendedor from SQLSERVER.desaerp.dbo.DIM_VENDEDORES where nombre='" & user & "'"
		set rs=oConn.execute(Sql) 
		nuser=rs.fields(0)
		sql="select top 1 numero_pedido FROM sqlserver.desaerp.dbo.PED_PEDIDOSENC as P where idvendedor=" & nuser & " and numero_pedido like '%" & nota & "' order by fecha_pedido DESC" 
		set rs=oConn.execute(Sql) 
		np=rs.fields(0)
		'exit sub
		else
			'call SinDatos()
			'exit sub
		end if
	end if
	if len(ao)=0 then ao=right(cstr(year(date-5)),2)

	'sql_cliente="SELECT c.RazonSocial, c.CodLegal, CAST(SUBSTRING(c.CtaCte, CHARINDEX('-', c.CtaCte) + 3, LEN(c.CtaCte) - CHARINDEX('-', c.CtaCte) + 3) AS numeric) AS local,  c.ListaPrecio, c.sigla, c.ejecutivo, p.* " & _
	'"FROM handheld.Flexline.FX_PEDIDO_PDA as P INNER JOIN Flexline.CtaCte  as C ON p.Cliente = c.CtaCte " & _
	'"WHERE (p.nota =  "& nota &") AND (p.vendedor = N'"& user &"') AND (c.Empresa = 'DESA') AND (c.TipoCtaCte = 'CLIENTE') and (left(p.fecha,2) = '"& ao &"') " & _
	'"ORDER BY  p.id desc"
	
	sql_cliente="SELECT c.RazonSocial, c.CodLegal, p.idsucursal AS local,  c.ListaPrecio, c.sigla, c.ejecutivo, p.* , c.ctacte as Cliente " & _
	"FROM sqlserver.desaerp.dbo.PED_PEDIDOSENC as P INNER JOIN Flexline.CtaCte  as C ON p.idcliente+' '+cast(p.idsucursal as nvarchar) = c.CtaCte " & _
	"WHERE (c.Empresa = '" & empresa & "') AND (c.TipoCtaCte = 'CLIENTE') and (p.numero_pedido = N'" & np & "')  " 
	if len(idempresa)>0 then sql_cliente=replace(sql_cliente,"and (p.numero_pedido = N"," and p.idempresa=" & idempresa & " and (p.numero_pedido = N")

	 sql_LineasProducto="SELECT L.*, P.Glosa " & _
	"FROM  sqlserver.desaerp.dbo.PED_PEDIDOSDET as L " & _
	" inner join serverdesa.BDFlexline.flexline.producto as P on P.producto=L.Producto " & _
	"WHERE     (L.numero_pedido = N'" & np & "') " & _
	"and p.empresa='" & empresa & "' " & _
	"ORDER BY L.linea_pedido"
	if len(idempresa)>0 then sql_LineasProducto=replace(sql_LineasProducto,"(L.numero_pedido = N","idempresa=" & idempresa & " and (L.numero_pedido = N")

	sql_notapedido="SELECT N.* " & _
	"FROM  sqlserver.desaerp.dbo.PED_PEDIDOSENC as N " & _
	"WHERE     (N.numero_pedido = N'" & np & "') "
	if len(idempresa)>0 then sql_notapedido=replace(sql_notapedido,"(N.numero_pedido = N"," idempresa=" & idempresa & " and (N.numero_pedido = N")

	set rs=oConn.execute(Sql)
	if rs.eof then 
		sql=replace(sql,"handheld.Flexline.FX_PEDIDO_PDA","handheld.dbo.DesaERP_PEDIDOS")
		set rs=oConn.execute(Sql)
	end if

	call NotaPedido()
End sub 'buscadatos()
'--------------------------------------------------------------------------------------------
Sub NotaPedido()
	%><html>
	<head>
	<title>Nota de Pedido</title>
	<style type="text/css">
	<!--
	body {
		background-color: #FFFFFF;
		font-family: Arial;
		font-size: 8;
	}
	td {
		font-size: 12px;
		border-collapse: collapse;
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
			<script type="text/javascript" language="javascript">
			
			if(1==<%=duplica%>)
			{
			alert("Pedido Duplicado");
			}
			
			function duplicaped()
			{
				window.location="./buscanota.asp?np=<%=np%>&idempresa=<%=idempresa%>&duplica=1";				
			}
			
			</script>
	<body>
	<CENTER>
	<TABLE width="650" name="marco"><TR><TD>
<!-- Datos Cliente -->
<% set rs=oConn.execute(sql_cliente)
if rs.eof then 
	 set rs=oConn.execute(replace(sql_cliente,"'08'","'09'"))
	 on error resume next
	 'response.write(sql_cliente)	 
end if
ejecutivo=rs("ejecutivo") %>
	<TABLE>
	<TR>
		<TD width="110"><B>Nota&nbsp;Pedido</B></TD>
		<TD onClick="window.open('detallera.asp?nota=<%=np%>')">:&nbsp;<B><%=np%></B></TD>
	</TR>
	<TR>
		<TD><B>Codigo&nbsp;Legal</B></TD>
		<TD>:&nbsp;<%=rs("codlegal")%>&nbsp;&nbsp;&nbsp;local&nbsp;:&nbsp;<B><%=rs("local")%></B></TD>
	</TR>
	<TR>
		<TD><B>Razon Social</B></TD>
		<TD>:&nbsp;<%=pc(rs("RazonSocial"))%>&nbsp;&nbsp;&nbsp;&nbsp;<B>Sigla</B>&nbsp;:&nbsp;<%=pc(rs("sigla"))%></TD>
	</TR>
	<%
	if idempresa=1 then empresa="DESA"
	if idempresa=4 then empresa="LACAV"
	if idempresa=3 then empresa="DESAZOFRI"
	
	Sql="SELECT * FROM serverdesa.BDFlexline.flexline.CtaCteDirecciones WHERE (CtaCte = '"& rs.fields("cliente") &"') AND "&_
	"(Principal <> 's') AND (TipoCtaCte = 'cliente') AND (Empresa = '" & empresa & "')"
	set rsd=oConn.execute(Sql)
	'response.write("empresa : " & idempresa & "<BR>")
	'response.write(sql)
	if not rsd.eof then 
		direccion=replace(rsd("direccion"),"|","-") 
		comuna=rsd("comuna") 
		ciudad=rsd("ciudad") 
	End if %>
	<TR>
		<TD><B>Direccion</B></TD>
		<TD>:&nbsp;<%=pc(direccion)%></TD>
	</TR>
	<TR>
		<TD><B>Comuna</B></TD>
		<TD>:&nbsp;<%=pc(comuna)%>&nbsp;&nbsp;&nbsp;&nbsp;<B>Ciudad:</B>&nbsp;<%=pc(ciudad)%></TD>
	</TR>

<!-- Datos Documento -->
<% set rs=oConn.execute(sql_notapedido) 
	if rs("IDEmpresa")="1" then nempresa="Distribucion Excelencia S.A."
	if rs("IDEmpresa")="4" then nempresa="Distribuidora La CAV"
	if rs("IDEmpresa")="3" then nempresa="Distribucion Excelencia (DesaZofri)"
	if rs("IDEmpresa")="1" then empresa="DESA"
	if rs("IDEmpresa")="4" then empresa="LACAV"
	if rs("IDEmpresa")="3" then empresa="DESAZOFRI"
	
	'idempresa = rs("IDEmpresa")
	clasevendedor="Cliente"
	if rs("IDEmpresa")="4" then clasevendedor="DESA"
	'if rs("IDEmpresa")="3" then clasevendedor="DESA"
	tipod =rs("tipodocto")
	sw_iva=rs("sw_iva"   )

	estado=ucase(rs("sw_estado"))
	financiero=ucase(rs("vb_aprfin"))
	stock=ucase(rs("vb_stock"))
	
	'response.write("DATOS PEDIDO : " & estado & financiero & stock & lenusped & "<BR>")
	
	IF (estado="R" AND financiero="S" AND stock="N" AND nuser<>"34" AND nuser<>"17" AND nuser<>"18" AND nuser<>"19" AND isNull(rs("usuario_pedido")) ) THEN
%>
	<TR>
		<td> 
			<BR><input type="BUTTON" value="Repetir Pedido" onclick="duplicaped()" style="font-size:10px" title="Repite pedido completo solo rechazado por quiebre">
		</td>
	</TR>
<% END IF
%>
</TABLE>
<HR>
	<TABLE>
	<TR>
		<TD width="110"><B>Estado</B></TD>
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
		<TD><B>Origen</B></TD>
		<TD>:&nbsp;<%
		if rs("Origen_notapedido")="O" then norigen="Oficina"
		if rs("Origen_notapedido")="M" then norigen="Movil"
		if rs("Origen_notapedido")="H" then norigen="HTC"
		response.write(rs("Origen_notapedido") & "&nbsp;:&nbsp;" & norigen)
		%></TD>
	</TR>
	<TR>
		<TD><B>Empresa</B></TD>
		<TD>:&nbsp;<%=rs("IDEmpresa")%>&nbsp;:&nbsp;<%=nempresa%></TD>
	</TR>
	<TR>
		<TD><B>Tipo Documento</B></TD>
		<TD>:&nbsp;<%=rs("tipodocto")%>&nbsp;:&nbsp;<B>sw_iva :</B> <%=rs("sw_iva")%></TD>
	</TR>
	<TR>
		<TD><B>Orden&nbsp;Compra</B></TD>
		<TD>:&nbsp;<%=rs("Orden_compra")%></TD>
	</TR>
	<TR>
<% If rs("correlativo_flex") > "0" Then %>	
		<TD><B>Correlativo Flex</B></TD>
		<TD>:&nbsp;<%=rs("correlativo_flex")%>&nbsp;>>&nbsp;<A HREF="documento.asp?empresa=<%=empresa%>&documento=<%=rs("correlativo_flex")%>">Ver&nbsp;Factura</A></TD>
<% Else %>
		<TD><B>Correlativo Factura</B></TD>
		<TD>:&nbsp;<%=rs("factura_sap")%>&nbsp;>>&nbsp;<A HREF="documento.asp?empresa=<%=empresa%>&documento=<%=rs("factura_sap")%>">Ver&nbsp;Factura</A></TD>		
<% End If %>
	</TR>
	<TR>
		<TD><B>Fecha&nbsp;Pedido</B></TD>
		<TD>:&nbsp;<%=ffecha(rs("fecha_pedido")) & "&nbsp;&nbsp;&nbsp;Ingreso&nbsp;:&nbsp;" & fhora(rs("hora_pedido")) & "&nbsp;&nbsp;&nbsp;Recepcion&nbsp;" & fhora(rs("Hora_recepcion")) %></TD>
	</TR>
	<TR>
		<TD><B>Fecha&nbsp;Proceso</B></TD>
		<TD>:&nbsp;<%=ffecha(rs("fecha_proceso")) %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B>Fecha Entrega</B>&nbsp;:&nbsp;<%=ffecha(rs("fecha_entrega")) %></TD>
	</TR>
	<TR>
		<TD><B>Fecha&nbsp;Facturacion</B></TD>
		<TD>:&nbsp;<%=ffecha(rs("fecha_Facturacion")) %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B>Fecha O/C</B>&nbsp;:&nbsp;<%=ffecha(rs("fecha_orden_compra")) %></TD>
	</TR>
	<TR>
		<TD><B>Usuario Aprobacion</B></TD>
		<TD>:&nbsp;<%=rs("Usuario_aprfin") %></TD>
	</TR>
	 	 Vb_aprfin
	<TR>
		<TD><B>Aprob. Automatica</B></TD>
		<TD>:&nbsp;<%=rs("Vb_apraut") %></TD>
	</TR>
	<TR>
		<TD><B>Aprob. Manual</B></TD>
		<TD>:&nbsp;<%=rs("Vb_aprfin") %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B>Credito Maximo</B>&nbsp;:&nbsp;<%=rs("Vb_cremax") %></TD>
	</TR>
	<TR>
		<TD><B>Deuda Ctacte</B></TD>
		<TD>:&nbsp;<%=rs("Vb_deucta") %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B>Protesto Vig.</B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp;<%=rs("Vb_provig") %></TD>
	</TR>
	<TR>
		<TD><B>Protesto Hist.</B></TD>
		<TD>:&nbsp;<%=rs("Vb_prohis") %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B>Dias Atra.</B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp;<%=rs("Vb_diaatr") %></TD>
	</TR>
	<TR>
		<TD><B>Descuento</B></TD>
		<TD>:&nbsp;<%=rs("Vb_precio") %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B>Stock</B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp;<%=rs("Vb_stock") %></TD>
	</TR>

	<TR>
		<TD><B>Vendedor Nota</B></TD>
		<TD>:&nbsp;<%=pc(user)%>&nbsp;&nbsp;&nbsp;<B>Vendedor&nbsp;<%=clasevendedor%>&nbsp;</B><%=pc(ejecutivo)%></TD>
	</TR>
	<TR>
		<TD><B>Local</B></TD>
		<%
		sql="select nombre from sqlserver.desaerp.dbo.dim_sucursales where idsucursal=" & rs("idlocal")
		nlocal=consultarapida(sql)

		sql="select nombre from sqlserver.desaerp.dbo.dim_bodegas where idbodega=" & rs("idbodega")
		nbodega=consultarapida(sql)
		%>
		<TD>:&nbsp;<%=pc(rs("idlocal") & ": " & nlocal)%>&nbsp;&nbsp;&nbsp;<B>Bodega&nbsp;</B><%=pc(rs("idbodega") & ": " & nbodega)%></TD>
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
	<TH>Desc.</TH>
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
	<TD align="center">-<%=rs("Descuento_Unitario"   )%>%</TD>
</TR>
<% rs.movenext
loop%>
</TABLE>
<!-- Totales -->
<%
':::: totales cantidad pedida ::::
'Reemplazar el valor de facped (Factor en lo pedido) por : 1

':::: totales cantidad despachada ::::
sql="Select T0.producto, T0.cantidad_pedida, T0.cantidad_despachada, T0.monto_flete, T0.monto_neto, T0.factor_ila,T1.fecha_facturacion " & _
"FROM sqlserver.desaerp.dbo.PED_PEDIDOSDET AS T0 INNER JOIN sqlserver.desaerp.dbo.PED_PEDIDOSENC AS T1 ON T0.numero_pedido=T1.numero_pedido and T0.idempresa=T1.idempresa " & _
"where T0.numero_pedido = '" & np & "' and T0.idempresa=" & idempresa & " "
'response.write(sql)
set rs=oConn.execute(sql)
serdis=0
afeiva=0
valiva=0
valtot=0
ilavin=0
ilacer=0
ilalic=0
ilawhi=0
ilabeb=0
Do until rs.eof
		'facped=cdbl(rs("cantidad_despachada"))/cdbl(rs("cantidad_pedida")) ':Cantidad Despachada
		facped=1 ':Cantidad Pedida
		serdis=serdis + (cdbl(rs("monto_flete")) * facped )
		afeiva=afeiva + (cdbl(rs("monto_neto")) * facped )
if rs("fecha_facturacion")<="20140930"then
		if rs("factor_ila")="02" then ilavin=ilavin+( cdbl(rs("Monto_Neto")) * 0.15 * facped )
		if rs("factor_ila")="15" then ilacer=ilacer+( cdbl(rs("Monto_Neto")) * 0.15 * facped )
		if rs("factor_ila")="03" then ilalic=ilalic+( cdbl(rs("Monto_Neto")) * 0.27 * facped )
		if rs("factor_ila")="16" then ilawhi=ilawhi+( cdbl(rs("Monto_Neto")) * 0.27 * facped )
		if rs("factor_ila")="18" then ilabeb=ilabeb+( cdbl(rs("Monto_Neto")) * 0.13 * facped )
else
sqlporcentaje="select AnalisisProducto7 as porcentaje " & _
"from serverdesa.BDFlexline.flexline.Producto " & _
"where producto = '" & rs("producto") & "' and empresa='DESA' "

set rsporcentaje=oConn.execute(sqlporcentaje)

		if left(rs("producto"),2)="VN" then ilavin=ilavin+( cdbl(rs("Monto_Neto")) * 0.205 * facped )
		if left(rs("producto"),2)="VI" then ilavin=ilavin+( cdbl(rs("Monto_Neto")) * 0.205 * facped )
		if left(rs("producto"),2)="CI" and left(rs("producto"),7)<>"CI37011" and left(rs("producto"),7)<>"CI37015" then ilacer=ilacer+( cdbl(rs("Monto_Neto")) * 0.205 * facped )
		if left(rs("producto"),7)="CI37011" then ilabeb=ilabeb+( cdbl(rs("Monto_Neto")) * 0.10 * facped )
		if left(rs("producto"),7)="CI37015" then ilabeb=ilabeb+( cdbl(rs("Monto_Neto")) * 0.10 * facped )
		if left(rs("producto"),2)="CN" then ilacer=ilacer+( cdbl(rs("Monto_Neto")) * 0.205 * facped )
		if left(rs("producto"),2)="LI" then ilalic=ilalic+( cdbl(rs("Monto_Neto")) * 0.315 * facped )
		if left(rs("producto"),2)="PN" then ilalic=ilalic+( cdbl(rs("Monto_Neto")) * 0.315 * facped )
		if left(rs("producto"),2)="PI" then ilalic=ilalic+( cdbl(rs("Monto_Neto")) * 0.315 * facped )
		if left(rs("producto"),2)="WH" then ilawhi=ilawhi+( cdbl(rs("Monto_Neto")) * 0.315 * facped )
		if left(rs("producto"),2)="BE" then ilabeb=ilabeb+( cdbl(rs("Monto_Neto")) * cdbl(rsporcentaje("porcentaje")/100) * facped )
end if
	rs.movenext
loop
afeiva=afeiva + serdis
valiva=afeiva*0.19
valtot=afeiva+valiva+ilavin+ilacer+ilalic+ilawhi+ilabeb+exeiva

'response.write(tipod)
if empresa<>"DESAZOFRI" then

	serdis=formatnumber(serdis,0)
	afeiva=formatnumber(afeiva,0)
	exeiva=formatnumber(exeiva,0)
	valiva=formatnumber(valiva,0)
	ilavin=formatnumber(ilavin,0)
	ilacer=formatnumber(ilacer,0)
	ilalic=formatnumber(ilalic,0)
	ilawhi=formatnumber(ilawhi,0)
	ilabeb=formatnumber(ilabeb,0)
	valtot=formatnumber(valtot,0)
	%>
	<p align="right">

	<TABLE width="150">
	<TR>
		<TD align="left" >Serv.Distri</TD>
		<TD align="right"><%=serdis%></TD>
	</TR>
	<TR>
		<TD align="left" >Afecto. IVA</TD>
		<TD align="right"><%=afeiva%></TD>
	</TR>
	<TR>
		<TD align="left" >Exento IVA</TD>
		<TD align="right"><%=exeiva%></TD>
	</TR>
	<TR>
		<TD align="left" >I.V.A.</TD>
		<TD align="right"><%=valiva%></TD>
	</TR>
	<TR>
		<TD align="left" >ILA Vinos</TD>
		<TD align="right"><%=ilavin%></TD>
	</TR>
	<TR>
		<TD align="left" >ILA Cerveza</TD>
		<TD align="right"><%=ilacer%></TD>
	</TR>
	<TR>
		<TD align="left" >ILA Licores</TD>
		<TD align="right"><%=ilalic%></TD>
	</TR>
	<TR>
		<TD align="left" >ILA Whisky</TD>
		<TD align="right"><%=ilawhi%></TD>
	</TR>
	<TR>
		<TD align="left" >ILA Bebida</TD>
		<TD align="right"><%=ilabeb%></TD>
	</TR>
	<TR>
		<TD align="left" ><B>Total</B></TD>
		<TD align="right"><B><%=valtot%></B></TD>
	</TR>
	</TABLE>Totales Segun cantidad Pedida
	<%
end if
%>

</p>
<!-- Observacion -->
<% sql="SELECT * FROM sqlserver.desaerp.dbo.PED_PEDIDOSOBS where numero_pedido='" & np & "' and idempresa=" & idempresa
set rs=oConn.execute(sql)
Response.write("<BR><B>Obs:&nbsp;</B>")
if not rs.eof then
	Response.write( trim(rs("observacion1")) )
End if
Response.write("<HR>")
%>

<!-- Etado Pedido --><BR>
<%
ap_fin="-No Hay"
ap_pre="-No Hay"
ap_sto="-No Hay"

sql="SELECT * " & _
	"FROM  sqlserver.desaerp.dbo.PED_PEDIDOSENC as N " & _
	"WHERE     (N.numero_pedido = N'" & np & "') and n.idempresa=" & idempresa & " " & _
	"and N.sw_estado in('F','R') and vb_aprfin='N'"
set rs=oConn.execute(sql)
if not rs.eof then 
	ap_fin="<TABLE>"
	if rs("vb_cremax")="N" then ap_fin=ap_fin & "<TR><TD>Credito Maximo Excedido</TD></TR>"
	if rs("vb_deucta")="N" then ap_fin=ap_fin & "<TR><TD>Atraso en Cuenta Corriente</TD></TR>"
	if rs("vb_provig")="N" then ap_fin=ap_fin & "<TR><TD>Protestos Vigentes</TD></TR>"
	if rs("vb_prohis")="N" then ap_fin=ap_fin & "<TR><TD>Protestos Historicos</TD></TR>"	
	ap_fin=ap_fin & "</TABLE>"
end if

sql="SELECT * " & _
	"FROM  sqlserver.desaerp.dbo.PED_PEDIDOSENC as N " & _
	"inner join sqlserver.desaerp.dbo.PED_PEDIDOSDET as L on N.Numero_pedido=L.Numero_pedido " & _
	"inner join serverdesa.BDFlexline.flexline.producto as P on L.producto=P.Producto " & _
	"inner join  sqlserver.desaerp.dbo.DIM_EMPRESAS as E on E.nombre=P.empresa and e.idempresa=N.idempresa " & _
	"WHERE     (N.numero_pedido = N'" & np & "')  and n.idempresa=" & idempresa & " " & _
	"and N.vb_aprfin='S' and N.sw_estado in('F','R') and L.vb_precio='N' " & _
	"order by L.linea_pedido"
	'response.write(sql)
set rs=oConn.execute(sql)
if not rs.eof then 
	ap_pre="<TABLE>"
	do until rs.eof
		ap_pre=ap_pre & "<TR><TD>" & rs("producto") & "</TD><TD>" & pc(rs("glosa")) & "</TD><TD>Descuento No Aprobado</TD></TR>"
	rs.movenext
	loop
	ap_pre=ap_pre & "</TABLE>"
End if

sql="SELECT *, L.cantidad_pedida-L.cantidad_despachada as Faltan " & _
	"FROM  sqlserver.desaerp.dbo.PED_PEDIDOSENC as N " & _
	"inner join sqlserver.desaerp.dbo.PED_PEDIDOSDET as L on N.Numero_pedido=L.Numero_pedido " & _
	"inner join serverdesa.BDFlexline.flexline.producto as P on L.producto=P.Producto " & _
	"inner join  sqlserver.desaerp.dbo.DIM_EMPRESAS as E on E.nombre=P.empresa and e.idempresa=N.idempresa " & _
	"WHERE     (N.numero_pedido = N'" & np & "')  and N.idempresa=" & idempresa & " " & _
	"and N.vb_aprfin='S' and N.sw_estado in('F','R')  and L.vb_precio='S' " & _
	"and (L.cantidad_pedida-L.cantidad_despachada)<>0 " & _
	"order by L.linea_pedido"
	'and L.vb_stock='N'
set rs=oConn.execute(sql)
if not rs.eof then 
	ap_sto="<TABLE>"
	do until rs.eof
		ap_sto=ap_sto & "<TR><TD>" & rs("producto") & "</TD><TD>" & pc(rs("glosa")) & "</TD><TD><B>faltan " & rs("faltan") & " unidades</B></TD></TR>"
	rs.movenext
	loop
	ap_sto=ap_sto & "</TABLE>"
End if

%>
<TABLE>
<TR>
	<TD width="20"></TD>
	<TD></TD>
	<TD></TD>
</TR>
<TR>
	<TD colspan="3"><B><U>Incidencias de Aprobaci&oacute;n</U></B></TD>
</TR>
<TR>
	<TD></TD>
	<TD><B>De tipo Financiero</B></TD>
	<TD></TD>
</TR>
<TR>
	<TD></TD>
	<TD><%=ap_fin%></TD>
	<TD></TD>
</TR>
<TR>
	<TD></TD>
	<TD><B>De tipo Precio</B></TD>
	<TD></TD>
</TR>
<TR>
	<TD></TD>
	<TD><%=ap_pre%></TD>
	<TD></TD>
</TR>
<TR>
	<TD></TD>
	<TD><B>De tipo Disponibilidad</B></TD>
	<TD></TD>
</TR>
<TR>
	<TD></TD>
	<TD><%=ap_sto%></TD>
	<TD></TD>
</TR>
</TABLE>
	</TD></TR></TABLE>
	</CENTER>
	</body>
	<%
End Sub 'NotaPedido()
'--------------------------------------------------------------------------------------------
function pc(texto)
texto=cisnull(texto," ")
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
'--------------------------------------------------------------------------------------------
Sub SinDatos()
	%><html>
	<head>
	<title>DESA</title>
	</head><body>
	<p align='center'>
	<BR><BR>Se necesitan datos
	<BR><input type="button" value="<< Atras" onClick="history.back()">
	</p>
	</body>
	</html><%
End Sub 'SinDatos()
'--------------------------------------------------------------------------------------------
function ffecha(fecha)
	dd=right(fecha,2)
	mm=left(right(fecha,4),2)
	yyyy=left(fecha,4)
	ffecha=dd & "/" & mm & "/" & yyyy
End function 'ffecha
'--------------------------------------------------------------------------------------------
function fhora(hora)
	hora=right("000000" & hora,6)
	hh=left(hora,2)
	mm=left(right(hora,4),2)
	ss=right(hora,2)
	fhora=hh & ":" & mm '& " :" & ss
End function 'fhora
'--------------------------------------------------------------------------------------------
Function cisnull(eval,altr)
	if isnull(eval) then
		cisnull=altr
	else
		cisnull=eval
	end if
End Function
'--------------------------------------------------------------------------------------------
Function consultarapida(mysql)
	set rcr=oConn.execute(mysql)
	if not rcr.eof then consultarapida=cisnull(rcr(0),"")
End Function
'--------------------------------------------------------------------------------------------
%>