<!--#include Virtual="/includes/migrilla_2.asp"-->
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
<HR>
<%
response.flush()
response.write("<CENTER><B>" & pc("Pedidos DesaERP",2) & "</B></CENTER>")

Pedidos_desaerp_sql=""& vbCrLf &_
"select '<A%20HREF=***../palm/buscanota.asp?np='                                               "& vbCrLf &_
"       + p.numero_pedido + '&idempresa='                                                      "& vbCrLf &_
"       + Cast(p.idempresa as nvarchar) + '***>'                                               "& vbCrLf &_
"       + right(p.numero_pedido, 4) + '</A>'                                    as Pedido,     "& vbCrLf &_
"       left(right('000000'+Cast(p.hora_pedido as nvarchar), 6), 2)                            "& vbCrLf &_
"       + ':'                                                                                  "& vbCrLf &_
"       + right(left(right('000000'+Cast(p.hora_pedido as nvarchar), 6), 4), 2) as Hora,       "& vbCrLf &_
"       '<B> ' + c.sigla + ' </B>'                                              as Sigla,      "& vbCrLf &_
"       left(c.razonsocial, 25)                                                 as RazonSocial,"& vbCrLf &_
"       '<B> ' + p.sw_estado + ' </B>'                                          as Est,        "& vbCrLf &_
"       '<A%20HREF=***/palm/documento.asp?idempresa=" & idempresa & "&factura='                "& vbCrLf &_
"       + (select top 1 nfactura                                                               "& vbCrLf &_
"          from   handheld.flexline.fx_pedido_pda as f                                         "& vbCrLf &_
"          where  nfactura not like 'X%'                                                       "& vbCrLf &_
"                 and notaventa = p.numero_pedido)                                             "& vbCrLf &_
"       + '***>'                                                                               "& vbCrLf &_
"       + isnull(Cast (p.factura_sap as nvarchar), Cast (p.factura_desa as nvarchar))          "& vbCrLf &_
"       + '</A>'                                                                as Detalle,    "& vbCrLf &_
"       isnull(p.pedido_externo, ' ')                                           as Externo,    "& vbCrLf &_
"       '<A%20HREF=***/palm/ruta.asp?factura='                                                 "& vbCrLf &_
"       + (select top 1 nfactura                                                               "& vbCrLf &_
"          from   handheld.flexline.fx_pedido_pda as f                                         "& vbCrLf &_
"          where  nfactura not like 'X%'                                                       "& vbCrLf &_
"                 and notaventa = p.numero_pedido)                                             "& vbCrLf &_
"       + '***>' + 'despacho' + '</A>'                                          as despacho    "& vbCrLf &_
"from   sqlserver.desaerp.dbo.ped_pedidosenc as p                                              "& vbCrLf &_
"       inner join sqlserver.desaerp.dbo.dim_empresas as e                                     "& vbCrLf &_
"               on p.idempresa = e.idempresa                                                   "& vbCrLf &_
"       inner join handheld.flexline.ctacte as c                                               "& vbCrLf &_
"               on c.empresa = e.nombre                                                        "& vbCrLf &_
"                  and c.ctacte = ( p.idcliente + ' '                                          "& vbCrLf &_
"                                   + Cast(p.idsucursal as nvarchar) )                         "& vbCrLf &_
"where  p.fecha_pedido = " & fecha & "                                                         "& vbCrLf &_
"       and p.idvendedor = " & nuser & "                                                       "& vbCrLf &_
"       and p.idempresa = " & idempresa & "                                                    "& vbCrLf &_
"order  by p.numero_pedido                                                                     "


Pedidos_desaerp_sql=""& vbCrLf &_
"select '<A%20HREF=***../palm/buscanota.asp?np='                                               "& vbCrLf &_
"       + p.numero_pedido + '&idempresa='                                                      "& vbCrLf &_
"       + Cast(p.idempresa as nvarchar) + '***>'                                               "& vbCrLf &_
"       + right(p.numero_pedido, 4) + '</A>'                                    as Pedido,     "& vbCrLf &_
"       left(right('000000'+Cast(p.hora_pedido as nvarchar), 6), 2)                            "& vbCrLf &_
"       + ':'                                                                                  "& vbCrLf &_
"       + right(left(right('000000'+Cast(p.hora_pedido as nvarchar), 6), 4), 2) as Hora,       "& vbCrLf &_
"       '<B> ' + c.sigla + ' </B>'                                              as Sigla,      "& vbCrLf &_
"       left(c.razonsocial, 25)                                                 as RazonSocial,"& vbCrLf &_
"       '<B> ' + p.sw_estado + ' </B>'                                          as Est,        "& vbCrLf &_
"       '<A%20HREF=***/palm/documento.asp?idempresa=" & idempresa & "&factura='                "& vbCrLf &_
"       + isnull(Cast (p.factura_sap as nvarchar), Cast (p.factura_desa as nvarchar))          "& vbCrLf &_
"       + '***>'                                                                               "& vbCrLf &_
"       + isnull(Cast (p.factura_sap as nvarchar), Cast (p.factura_desa as nvarchar))          "& vbCrLf &_
"       + '</A>'                                                                as Detalle,    "& vbCrLf &_
"       isnull(p.pedido_externo, ' ')                                           as Externo,    "& vbCrLf &_
"       '<A%20HREF=***/palm/ruta.asp?factura='                                                 "& vbCrLf &_
"       + (select top 1 nfactura                                                               "& vbCrLf &_
"          from   handheld.flexline.fx_pedido_pda as f                                         "& vbCrLf &_
"          where  nfactura not like 'X%'                                                       "& vbCrLf &_
"                 and notaventa = p.numero_pedido)                                             "& vbCrLf &_
"       + '***>' + 'despacho' + '</A>'                                          as despacho    "& vbCrLf &_
"from   sqlserver.desaerp.dbo.ped_pedidosenc as p                                              "& vbCrLf &_
"       inner join sqlserver.desaerp.dbo.dim_empresas as e                                     "& vbCrLf &_
"               on p.idempresa = e.idempresa                                                   "& vbCrLf &_
"       inner join handheld.flexline.ctacte as c                                               "& vbCrLf &_
"               on c.empresa = e.nombre                                                        "& vbCrLf &_
"                  and c.ctacte = ( p.idcliente + ' '                                          "& vbCrLf &_
"                                   + Cast(p.idsucursal as nvarchar) )                         "& vbCrLf &_
"where  p.fecha_pedido = " & fecha & "                                                         "& vbCrLf &_
"       and p.idvendedor = " & nuser & "                                                       "& vbCrLf &_
"       and p.idempresa = " & idempresa & "                                                    "& vbCrLf &_
"order  by p.numero_pedido                                                                     "


Pedidos_desaerp_sql_2=""& vbCrLf &_
"select ''                                                                                     "& vbCrLf &_
"       + right(p.numero_pedido, 4) + ''                                    as Pedido,         "& vbCrLf &_
"       left(right('000000'+Cast(p.hora_pedido as nvarchar), 6), 2)                            "& vbCrLf &_
"       + ':'                                                                                  "& vbCrLf &_
"       + right(left(right('000000'+Cast(p.hora_pedido as nvarchar), 6), 4), 2) as Hora,       "& vbCrLf &_
"       ' ' + c.sigla + ' '                                              as Sigla,             "& vbCrLf &_
"       left(c.razonsocial, 25)                                                 as RazonSocial,"& vbCrLf &_
"       ' ' + p.sw_estado + ' '                                          as Est,               "& vbCrLf &_
"       ''                                                                                     "& vbCrLf &_
"       + isnull(Cast (p.factura_sap as nvarchar), Cast (p.factura_desa as nvarchar))          "& vbCrLf &_
"       + ''                                                                as Detalle,        "& vbCrLf &_
"       isnull(p.pedido_externo, ' ')                                           as Externo,    "& vbCrLf &_
"       '' + 'despacho' + ''                                          as despacho              "& vbCrLf &_
"from   sqlserver.desaerp.dbo.ped_pedidosenc as p                                              "& vbCrLf &_
"       inner join sqlserver.desaerp.dbo.dim_empresas as e                                     "& vbCrLf &_
"               on p.idempresa = e.idempresa                                                   "& vbCrLf &_
"       inner join handheld.flexline.ctacte as c                                               "& vbCrLf &_
"               on c.empresa = e.nombre                                                        "& vbCrLf &_
"                  and c.ctacte = ( p.idcliente + ' '                                          "& vbCrLf &_
"                                   + Cast(p.idsucursal as nvarchar) )                         "& vbCrLf &_
"where  p.fecha_pedido = 20170111                                                              "& vbCrLf &_
"       and p.idvendedor = 89                                                                  "& vbCrLf &_
"       and p.idempresa = 1                                                                    "& vbCrLf &_
"order  by p.numero_pedido                                                                     "

call migrilla(Pedidos_desaerp_sql, 2, 0,"","")

'/*	DEBUG */>>
	dev_sql = "<hr/><pre>"&Pedidos_desaerp_sql&"</pre><hr/>"
	response.write(dev_sql) ' solo comentar esto para probar
'/*	DEBUG */<<
response.flush()
response.write("<HR><CENTER><B>" & pc("Pedidos Por Call Center",2) & "</B></CENTER>")


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
