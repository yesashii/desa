<!--#include Virtual="/includes/migrilla.asp"-->
<!--#include Virtual="/includes/pda.asp"-->
<BODY topmargin="0">
<CENTER>
<%
nuser =trim(cstr(request.form("nuser")))
if len(nuser)=0 then nuser=request.querystring("nuser")
nuser=cint(nuser)
fecha =trim(cstr(request.form("mifecha")))
if len(fecha)=0 then fecha=request.querystring("mifecha")
if len(fecha)=0 then fecha=year(date) & right("00" & month(date),2) & right("00" & day(date),2)
	call encabezado(nuser,"")
fecha=cdbl(fecha)

empresa=trim(cstr(request.form("empresa")))
if len(empresa)=0 then empresa=request.querystring("empresa")

'response.write(empresa)

if empresa="DESAZOFRI" then idempresa=3
if empresa="DESA" then idempresa=1
if empresa="LACAV" then idempresa=4

'response.write("<HR>" & empresa & "<HR>")

sql="select nombre from SQLSERVER.desaerp.dbo.DIM_VENDEDORES where idvendedor=" & nuser
if len (empresa)>0 then sql="select nombre from SQLSERVER.desaerp.dbo.DIM_VENDEDORES where idvendedor=" & nuser & " and idempresa='" & idempresa & "'"
'response.write(sql)
set rs=oConn.execute(Sql) 
user  = rs.fields(0)

'response.write(user)

if fecha<20080220 then
%><FORM METHOD=POST ACTION="estadosped.asp">
<BR>Por Cambio de sistema ...<BR><BR>
<INPUT TYPE="hidden" name="midia" value="18">
<INPUT TYPE="hidden" name="mimes" value="02">
<INPUT TYPE="hidden" name="miano" value="08">
<INPUT TYPE="hidden" name="nuser" value="<%=right("000" & nuser,3)%>">
<INPUT TYPE="submit" value="Ir al antiguo reporte">
</FORM><%
End if
%>
<FORM METHOD=POST ACTION="">
<TABLE>
<TR>
	<TD>Fecha</TD>
	<TD>
	<SELECT NAME="mifecha">
	<%
	for x=0 to 40
	xselec=""
	rfecha=cdate(left(fecha,4) & "-" & mid(fecha,5,2) & "-" & right(fecha,2) ) -20+x
	xfecha=year(rfecha) & right("00" & month(rfecha),2) &  right("00" & day(rfecha),2)
	'if cint(right(rfecha,2))<1 then xfecha=rfecha-69 
	'if cint(right(rfecha,2))>31 then xfecha=rfecha+69
	if cdbl(fecha)=cdbl(xfecha) then xselec=" selected "
	response.write("<OPTION " & xselec & "value=" & xfecha & ">" & rfecha & "</OPTION>")
	'right(xfecha,2) & "/" & mid(xfecha,5,2) & "/" & left(xfecha,4
		if rfecha>(date()+2) then exit for
	next
	%>
	</SELECT>
	</TD>
	<TD><INPUT TYPE="hidden" name="nuser" value="<%=nuser%>"><INPUT TYPE="submit" value="Ver"></TD>
</TR>
</TABLE>
</FORM>
<FONT SIZE="1" face="verdana" COLOR="#CC0000">Estados I=No Procesado ; P=En Proceso ; F=Facturado ; R=Rechazado ; N=Nulo</FONT>
<HR><%
response.flush()
'http://pda.desa.cl/palm/documento.asp?factura=0000638188
'<A%20HREF=***/palm/documento.asp?factura='+pda.NFactura+'***>'+pda.NFactura+'</A>'
response.write("<CENTER><B>" & pc("Pedidos DesaERP",2) & "</B></CENTER>")
sql="SELECT " & _
"'<A%20HREF=***../palm/buscanota.asp?np='+pd.numero_pedido+'&idempresa='+pd.idempresa+'***>'+cast(pd.nota as nvarchar)+'</A>' as Pedido, " & _     
"	pd.Nombre_cliente, pd.Est, " & _
"'<A%20HREF=***/palm/documento.asp?factura='+NF.NFactura+'***>'+right(NF.NFactura,6)+'</A>' as Detalle " & _
"FROM         (SELECT  P.numero_pedido, RIGHT(P.numero_pedido, 4) AS Nota, LEFT('<B> ' + C.Sigla + '</B>' + C.RazonSocial, 40) AS Nombre_cliente, " & _
"                                              '<B> '+P.sw_estado+'</B>' AS Est , cast(p.idempresa as nvarchar) as idempresa " & _
"                       FROM          SQLSERVER.desaerp.dbo.PED_PEDIDOSENC AS P INNER JOIN  " & _
"                                              handheld.Flexline.CtaCte AS C ON P.idcliente + ' ' + CAST(P.idsucursal AS nvarchar) = C.CtaCte " & _
"                       WHERE      (P.idvendedor = " & nuser & ") AND (P.fecha_pedido = " & fecha & ") AND (C.Empresa = 'desa') AND (C.TipoCtaCte = 'cliente')) AS pd LEFT OUTER JOIN " & _
"                          (SELECT     LEFT(P.fecha, 4) + RIGHT('000' + CAST(V.idvendedor AS nvarchar), 3) + RIGHT('0000' + CAST(P.nota AS nvarchar), 4) AS Numero_pedido,  " & _
"                                                   P.NFactura " & _
"                            FROM          handheld.Flexline.FX_PEDIDO_PDA AS P INNER JOIN " & _
"                                                   SQLSERVER.desaerp.dbo.DIM_VENDEDORES AS V ON P.vendedor = V.nombre " & _
"                            WHERE      v.idempresa=1 and (P.NFactura NOT LIKE 'X%') AND (V.idvendedor = " & nuser & ")) AS NF ON pd.numero_pedido = NF.Numero_pedido " & _
"order by pd.Nota"

sql="SELECT " & _
"'<A%20HREF=***../palm/buscanota.asp?np='+pd.numero_pedido+'&idempresa='+pd.idempresa+'***>'+cast(pd.nota as nvarchar)+'</A>' as Pedido, " & _     
"	pd.Nombre_cliente, pd.Est, " & _
"'<A%20HREF=***/palm/documento.asp?factura='+NF.NFactura+'&idempresa=" & idempresa & "***>'+right(NF.NFactura,7)+'</A>' as Detalle " & _
"FROM         (SELECT  P.numero_pedido, RIGHT(P.numero_pedido, 4) AS Nota, LEFT('<B> ' + C.Sigla + '</B>' + C.RazonSocial, 40) AS Nombre_cliente, " & _
"                                              '<B> '+P.sw_estado+'</B>' AS Est, cast(p.idempresa as nvarchar) as idempresa " & _
"                       FROM          SQLSERVER.desaerp.dbo.PED_PEDIDOSENC AS P INNER JOIN  " & _
"                                              handheld.Flexline.CtaCte AS C ON P.idcliente + ' ' + CAST(P.idsucursal AS nvarchar) = C.CtaCte " & _
"                       WHERE      (P.idvendedor = " & nuser & ") AND (P.fecha_pedido = " & fecha & ") AND (C.Empresa in('DESA','DESAZOFRI')) AND (C.TipoCtaCte = 'cliente')) AS pd LEFT OUTER JOIN " & _
"                          (SELECT     P.notaventa AS Numero_pedido,  " & _
"                                                   P.NFactura, p.empresa " & _
"                            FROM          handheld.Flexline.FX_PEDIDO_PDA AS P INNER JOIN " & _
"                                                   SQLSERVER.desaerp.dbo.DIM_VENDEDORES AS V ON P.vendedor = V.nombre " & _
"                            WHERE      v.idempresa=1 and (P.NFactura NOT LIKE 'X%') AND (V.idvendedor = " & nuser & ")) AS NF " & _
"ON cast(pd.idempresa as nvarchar) + pd.numero_pedido = replace(replace(replace(NF.empresa,'DESA','1'),'LACAV','4'),'DESAZOFRI','3') + NF.Numero_pedido " & _
"group by pd.numero_pedido, pd.nota, pd.Nombre_cliente, pd.Est, NF.NFactura, pd.idempresa  " & _
"order by pd.Nota"


sqliqq="SELECT " & _
"'<A%20HREF=***../palm/buscanota.asp?np='+pd.numero_pedido+'&idempresa='+pd.idempresa+'***>'+cast(pd.nota as nvarchar)+'</A>' as Pedido, " & _     
"	pd.Nombre_cliente, pd.Est, " & _
"'<A%20HREF=***/palm/documento.asp?factura='+NF.NFactura+'***>'+right(NF.NFactura,7)+'</A>' as Detalle " & _
", pd.externo " & _
"FROM         (SELECT  P.numero_pedido, RIGHT(P.numero_pedido, 4) AS Nota, LEFT('<B> ' + C.Sigla + '</B>' + C.RazonSocial, 40) AS Nombre_cliente, " & _
"                                              '<B> '+P.sw_estado+'</B>' AS Est, cast(p.idempresa as nvarchar) as idempresa , p.pedido_externo as externo " & _
"                       FROM          SQLSERVER.desaerp.dbo.PED_PEDIDOSENC AS P INNER JOIN  " & _
"                                              handheld.Flexline.CtaCte AS C ON P.idcliente + ' ' + CAST(P.idsucursal AS nvarchar) = C.CtaCte " & _
"                       WHERE      (P.idvendedor = " & nuser & ") AND (P.fecha_pedido = " & fecha & ") AND (C.Empresa in('" & empresa & "')) AND (C.TipoCtaCte = 'cliente')) AS pd LEFT OUTER JOIN " & _
"                          (SELECT     P.notaventa AS Numero_pedido,  " & _
"                                                   P.NFactura, p.empresa " & _
"                            FROM          handheld.Flexline.FX_PEDIDO_PDA AS P INNER JOIN " & _
"                                                   SQLSERVER.desaerp.dbo.DIM_VENDEDORES AS V ON P.vendedor = V.nombre " & _
"                            WHERE      v.idempresa=1 and (P.NFactura NOT LIKE 'X%') AND (V.idvendedor = " & nuser & ")) AS NF " & _
"ON cast(pd.idempresa as nvarchar) + pd.numero_pedido = replace(replace(replace(NF.empresa,'DESA','1'),'LACAV','4'),'DESAZOFRI','3') + NF.Numero_pedido " & _
"group by pd.numero_pedido, pd.nota, pd.Nombre_cliente, pd.Est, NF.NFactura, pd.idempresa, pd.externo  " & _
"order by pd.Nota"

sql="select '<A%20HREF=***../palm/buscanota.asp?np='+p.numero_pedido+'&idempresa='+cast(p.idempresa as nvarchar)+'***>'+right(p.numero_pedido,4)+'</A>' as Pedido, " & _
"left(right('000000'+cast(p.Hora_pedido as nvarchar),6),2)+':'+right(left(right('000000'+cast(p.Hora_pedido as nvarchar),6),4),2) as Hora, " & _
"'<B> '+C.Sigla+' </B>' as Sigla, " & _
"left(C.RazonSocial,25) as RazonSocial, " & _
"'<B> '+P.sw_estado+' </B>' AS Est, " & _
"'<A%20HREF=***/palm/documento.asp?idempresa=" & idempresa & "&factura='+(select top 1 nfactura from handheld.Flexline.FX_PEDIDO_PDA as F where nfactura not like 'X%' and notaventa=p.numero_pedido )+'***>'+(  select top 1 nfactura " & _
"	from handheld.Flexline.FX_PEDIDO_PDA as F  " & _
"	where nfactura not like 'X%' and notaventa=p.numero_pedido " & _
")+'</A>' as Detalle, " & _
"isnull(p.pedido_externo,' ') as Externo " & _
",'<A%20HREF=***/palm/ruta.asp?factura='+(select top 1 nfactura from handheld.Flexline.FX_PEDIDO_PDA as F where nfactura not like 'X%' and notaventa=p.numero_pedido )+'***>'+'despacho'+'</A>' as despacho " & _
"from SQLSERVER.desaerp.dbo.PED_PEDIDOSENC as P inner join  " & _
"	SQLSERVER.desaerp.dbo.DIM_Empresas as E " & _
"	on P.idempresa=E.idempresa inner join  " & _
"	handheld.flexline.ctacte as C  " & _
"	on C.empresa=E.nombre and c.ctacte=(P.idcliente+' '+cast(P.idsucursal as nvarchar)) " & _
"where P.fecha_pedido = " & fecha & " and p.idvendedor=" & nuser & " and p.idempresa=" & idempresa & " " & _
"order by p.numero_pedido"

'"left(right('000000'+cast(p.Hora_pedido as nvarchar),6),2)+':'+right(left(right('000000'+cast(p.Hora_pedido as nvarchar),6),4),2) as Hora, " & _

'(select top 1 nfactura from handheld.Flexline.FX_PEDIDO_PDA as F where nfactura not like 'X%' and notaventa=p.numero_pedido )

'if empresa="DESAZOFRI" then sql=sqliqq
'sql=sqliqq
'response.write("<HR>" & sql & "<HR>")

call migrilla(SQL, 2, 0,"","")
'response.write(SQL)


'sql="SELECT '<A%20HREF=***../palm/buscanota.asp?user='+replace(DIM_vendedores_1.nombre,' ','%20')+'&nota='+cast(cast(Pedido.Pedido as numeric) as nvarchar)+'***>'+cast(Pedido.Pedido as nvarchar)+'</A>' as Pedido,left('<B>'+cta.sigla+'</B>&nbsp;' + Cta.RazonSocial,40) as 'Nombre Cliente', Pedido.Est, '<A%20HREF=***/palm/documento.asp?factura='+pda.NFactura+'***>'+right(pda.NFactura,6)+'</A>' as Detalle " & _
'"FROM         (SELECT     RIGHT(numero_pedido, 4) AS Pedido, CAST(idcliente AS nvarchar) + ' ' + CAST(idsucursal AS nvarchar) AS Cliente,  " & _
'"REPLACE(REPLACE(REPLACE(REPLACE(vb_aprfin + sw_estado, 'SP', ''), 'SF', 'OK'), 'SR', '<B>RD</B>'), 'NR', '<B>RD</B>') AS Est, idvendedor " & _
'"                       FROM          SQLSERVER.Desaerp.dbo.PED_PEDIDOSENC AS p " & _
'"                       WHERE      (fecha_pedido = " & fecha & ") AND (idvendedor = " & nuser & ")) AS Pedido INNER 'JOIN " & _
'"                      handheld.Flexline.CtaCte AS Cta ON Pedido.Cliente = Cta.CtaCte INNER JOIN " & _
'"                      handheld.Flexline.FX_PEDIDO_PDA AS pda ON Pedido.Pedido = pda.nota INNER JOIN " & _
'"                      SQLSERVER.Desaerp.dbo.DIM_vendedores AS DIM_vendedores_1 ON pda.vendedor = DIM_vendedores_1.nombre AND  " & _
'"                      Pedido.idvendedor = DIM_vendedores_1.idvendedor " & _
'"WHERE     (Cta.Empresa = 'desa') AND (Cta.TipoCtaCte = 'cliente') AND (LEFT(pda.fechaentrega, 4) = '" & left(fecha,4) & "' ) " & _
'"ORDER BY pedido.pedido"
'response.write(sql)
'	call migrilla(SQL, 2, 0,"","")
'response.write("<HR><CENTER><B>" & pc("Pedidos Sin Procesar",2) & "</B></CENTER>")

'sql="SELECT '<A%20HREF=***../palm/buscanota.asp?user='+replace(pda.vendedor,' ','%20')+'&nota='+cast(cast(pda.nota as numeric) as nvarchar)+'***>'+cast(pda.nota as nvarchar)+'</A>' " & _
'"AS Pedido, LEFT('<B>' + Cta.Sigla + '</b>' + Cta.RazonSocial, 40) AS RazonSocial, pda.estado AS est, " & _
'"                      '' AS Detalle " & _
'"FROM         handheld.Flexline.FX_PEDIDO_PDA AS pda INNER JOIN " & _
'"                      handheld.Flexline.FX_vende_man as ven ON pda.vendedor = ven.nombre INNER JOIN " & _
'"                      handheld.Flexline.CtaCte as Cta ON pda.Cliente = Cta.CtaCte " & _
'"WHERE     (pda.estado = N'np') AND (pda.fecha = N'" & right(fecha,6) & "') AND (ven.numero = " & nuser & ") AND (Cta.Empresa = 'desa') AND (Cta.TipoCtaCte = 'cliente')"
'	call migrilla(SQL, 2, 0,"","")
'response.write( right(fecha,6) )




response.flush()
response.write("<HR><CENTER><B>" & pc("Pedidos Por Call Center",2) & "</B></CENTER>")







'sql="declare @nombrevendedor as nvarchar(50) " & _
'"set @nombrevendedor=(select nombre from sqlserver.desaerp.dbo.DIM_VENDEDORES where idvendedor=" & nuser & ") " & _
'"SELECT     TOP (100) PERCENT '<A%20HREF=***../palm/buscanota.asp?user=' + REPLACE(DIM_vendedores_1.nombre, ' ', ''%20') " & _ 
'"+ '&nota=' + CAST(CAST(Pedido.Pedido AS numeric) AS nvarchar) + '***>' + CAST(Pedido.Pedido AS nvarchar) + '</A>' AS Pedido, " & _ 
'" LEFT('<B>' + Cta.Sigla + '</B>&nbsp;' + Cta.RazonSocial, 40) AS 'Nombre Cliente', Pedido.Est,  " & _
'" '<A%20HREF=***/palm/documento.asp?factura=' + pda.NFactura + '***>' + RIGHT(pda.NFactura, 6) + '</A>' AS Detalle " & '_
'"FROM         (SELECT     RIGHT(numero_pedido, 4) AS Pedido, CAST(idcliente AS nvarchar) + ' ' + CAST(idsucursal AS nvarchar) AS Cliente,  " & _
'" REPLACE(REPLACE(REPLACE(vb_aprfin + sw_estado, 'SF', 'OK'), 'SR', '<B>RD</B>'), 'NR', '<B>RD</B>') AS Est, idvendedor  " & _
'" FROM          SQLSERVER.Desaerp.dbo.PED_PEDIDOSENC AS p  " & _
'" WHERE      (fecha_pedido = " & fecha & ") AND (idvendedor IN (17, 18, 19))) AS Pedido INNER JOIN " & _
'" handheld.Flexline.CtaCte AS Cta ON Pedido.Cliente = Cta.CtaCte INNER JOIN " & _
'" handheld.Flexline.FX_PEDIDO_PDA AS pda ON Pedido.Pedido = pda.nota INNER JOIN " & _
'" SQLSERVER.Desaerp.dbo.DIM_vendedores AS DIM_vendedores_1 ON pda.vendedor = DIM_vendedores_1.nombre AND  " & _
'" Pedido.idvendedor = DIM_vendedores_1.idvendedor " & _
'"WHERE     (Cta.Empresa = 'desa') AND (Cta.TipoCtaCte = 'cliente') AND (Cta.Ejecutivo = @nombrevendedor ) " & _
'"ORDER BY Pedido"
'	call migrilla(SQL, 2, 0,"","")
sql="SELECT " & _
"'<A%20HREF=***../palm/buscanota.asp?np='+pd.numero_pedido+'***>'+cast(pd.nota as nvarchar)+'</A>' as Pedido, " & _     
"	pd.Nombre_cliente, pd.Est, " & _
" CASE len(NF.NFactura) when 0 then '' else " & _
"'<A%20HREF=***/palm/documento.asp?factura='+NF.NFactura+'***>'+right(NF.NFactura,7)+'</A>' " & _
" END as Detalle " & _
"FROM (SELECT P.numero_pedido, RIGHT(P.numero_pedido, 4) AS Nota, " & _
"LEFT('<B> ' + C.Sigla + '</B>' + C.RazonSocial, 40) AS Nombre_cliente, " & _
" '<B> '+P.sw_estado+'</B>' AS Est , p.correlativo_flex " & _
" FROM SQLSERVER.desaerp.dbo.PED_PEDIDOSENC AS P INNER JOIN  " & _
" handheld.Flexline.CtaCte AS C ON P.idcliente + ' ' + CAST(P.idsucursal AS nvarchar) = C.CtaCte " & _
" WHERE (c.ejecutivo = '" & user & "') and (p.idvendedor<>" & nuser & ") AND (P.fecha_pedido = " & fecha & ") AND (C.Empresa = 'desa') " & _
" AND (C.TipoCtaCte = 'cliente') and p.idempresa=1) AS pd LEFT OUTER JOIN " & _
" (SELECT p.notaventa AS Numero_pedido,  " & _
" isnull(P.NFactura,'') as NFactura " & _
" FROM handheld.Flexline.FX_PEDIDO_PDA AS P INNER JOIN " & _
" SQLSERVER.desaerp.dbo.DIM_VENDEDORES AS V ON P.vendedor = V.nombre " & _
" WHERE (P.NFactura NOT LIKE 'X%') AND (V.idvendedor in(17,18,19,300) ) " & _
" group by p.notaventa , P.NFactura " & _
") AS NF ON pd.numero_pedido = NF.Numero_pedido " & _
"order by pd.Nota"

'response.write("<BR>" & sql & "<BR>")


call migrilla(sql, 2, 0,"","")
'response.write(sql)
'----------------------------------------------------------------------------
function nombrevendedor(numero)

End function 'nombrevendedor
'----------------------------------------------------------------------------
%>
<TABLE>
<TR>
	<TD Colspan="2">Estados</TD>
</TR>
<TR>
	<TD><B>I&nbsp;:</B></TD>
	<TD>No Procesado</TD>
</TR>
<TR>
	<TD><B>P&nbsp;:</B></TD>
	<TD>En Proceso</TD>
</TR>
<TR>
	<TD><B>F&nbsp;:</B></TD>
	<TD>Facturado</TD>
</TR>
<TR>
	<TD><B>R&nbsp;:</B></TD>
	<TD>Rechazado</TD>
</TR>
<TR>
	<TD><B>N&nbsp;:</B></TD>
	<TD>Nulo</TD>
</TR>
</TABLE>
</CENTER>
</BODY>
