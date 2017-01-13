<body bgcolor="#FFFFFF" topmargin="0">
<%
dim factura, mibusca
factura=request.querystring("factura")
if len(factura)=0 then factura =trim(cstr(request.form("factura")))
MiBusca=trim(cstr(request.form("mibusca")))
'response.write(factura)
if len(factura) < 4 then 
	if len(mibusca)<1 then
	call nuevabus()
	else
	call buscapor()
	end if
else
	call buscafac()
	'call DatosFactuRA()
end if

'------------------------------------------------------------------------------------------------------
':::::Pie de pagina::::::::::
private sub nuevabus()
%>
<CENTER>
<FONT SIZE="1" COLOR="#000099">Busca documentos por su numero</FONT>
<BR>
<FORM METHOD=POST ACTION="Documento.asp">
<INPUT TYPE="text" NAME="factura" size="10">
<INPUT TYPE="submit" value="Buscar">
</FORM>
</CENTER>
<center>
<HR>
<FONT SIZE="1" COLOR="#000099">Busca documentos por Nombre o Rut</FONT>
<FORM METHOD=POST ACTION="Documento.asp">
<TABLE>
<TR>
	<TD><INPUT TYPE="text" NAME="mibusca"></TD>
</TR>
<TR>
	<TD><INPUT TYPE="radio" NAME="buspor" value="rs">Razon Social</TD>
</TR>
<TR>
	<TD><INPUT TYPE="radio" NAME="buspor" value="rut" CHECKED>Codigo legal</TD>
</TR>
<TR>
	<TD><INPUT TYPE="submit" value="Buscar"></TD>
</TR>

</TABLE>
</FORM>
<HR>
<form method=post action=menu.asp>
<input type=submit value="<< Menu">
</form>
<%
end sub 'nuevabus()
'------------------------------------------------------------------------------------------------------
private sub buscafac()
'''''''''''if len(factura)<10 then factura=right("0000000000" & factura,10)

NotaVenta = factura
':::::::::::::::::: conexion :::::::::::::::::
Dim tipodoc, mitotal, oConn, rs, sql, oconn1, rs1
Set oConn = server.createobject("ADODB.Connection")
Set oConn1 = server.createobject("ADODB.Connection")
oConn.open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=flexline2;User Id=sa;Password=desakey;"

oConn1.open "Provider=SQLOLEDB;Data Source=serverdesa;Initial Catalog=BDFlexline;User Id=sa;Password=desakey;"

'"FROM `W:\TRASPASOSPALM`.Notas_ventas Notas_ventas, `W:\TRASPASOSPALM`.Notas_ventas_detalles Notas_ventas_detalles " & _

'SQL="SELECT TipoDocumento.TipoDocto, Documento.Numero, Documento.fecha, CtaCte.CtaCte, CtaCte.RazonSocial, CtaCteDirecciones.Direccion, DocumentoD.Producto, PRODUCTO.GLOSA, DocumentoD.Cantidad, DocumentoD.Neto, DocumentoD.PorcentajeDR, Documento.total, Documento.Vigencia, ctacte.comuna, ctacte.ejecutivo " & _
'"FROM BDFlexline.flexline.CtaCte CtaCte, BDFlexline.flexline.CtaCteDirecciones CtaCteDirecciones, BDFlexline.flexline.Documento Documento, BDFlexline.flexline.DocumentoD DocumentoD, BDFlexline.flexline.PRODUCTO PRODUCTO, BDFlexline.flexline.TipoDocumento TipoDocumento " & _
'"WHERE CtaCteDirecciones.CtaCte = CtaCte.CtaCte AND CtaCteDirecciones.Empresa = CtaCte.Empresa AND CtaCteDirecciones.TipoCtaCte = CtaCte.TipoCtaCte AND Documento.Empresa = CtaCte.Empresa AND Documento.Empresa = CtaCteDirecciones.Empresa AND Documento.TipoCtaCte = CtaCte.TipoCtaCte AND DocumentoD.Correlativo = Documento.Correlativo AND DocumentoD.Empresa = CtaCte.Empresa AND DocumentoD.Empresa = CtaCteDirecciones.Empresa AND DocumentoD.Empresa = Documento.Empresa AND DocumentoD.TipoDocto = Documento.TipoDocto AND CtaCte.CtaCte = Documento.Cliente AND PRODUCTO.EMPRESA = CtaCte.Empresa AND PRODUCTO.EMPRESA = Documento.Empresa AND PRODUCTO.EMPRESA = DocumentoD.Empresa AND DocumentoD.Producto = PRODUCTO.PRODUCTO AND TipoDocumento.Empresa = CtaCte.Empresa AND TipoDocumento.Empresa = CtaCteDirecciones.Empresa AND TipoDocumento.Empresa = Documento.Empresa AND TipoDocumento.Empresa = DocumentoD.Empresa AND TipoDocumento.Empresa = PRODUCTO.EMPRESA AND TipoDocumento.TipoDocto = Documento.TipoDocto AND TipoDocumento.TipoDocto = DocumentoD.TipoDocto AND TipoDocumento.TipoCtaCte = CtaCte.TipoCtaCte AND ((CtaCte.Empresa='desa') AND (CtaCte.TipoCtaCte='CLIENTE') AND (CtaCteDirecciones.Principal<>'s') AND (Documento.Numero='" & factura & "') AND (TipoDocumento.Clase='Factura (v)') AND (TipoDocumento.Factormonto<>0))"

SQL="SELECT Notas_ventas.Nota_venta, Notas_ventas.Vendedor_empresa, Notas_ventas.Fecha_nota_venta, Notas_ventas.Hora_recepcion, Notas_ventas.Rut_cliente, Notas_ventas.Direccion_despacho, Notas_ventas.Forma_pago, Notas_ventas_detalles.Linea, Notas_ventas_detalles.Producto, Notas_ventas_detalles.Cantidad_original, Notas_ventas_detalles.Precio, Notas_ventas_detalles.Descuento2, Notas_ventas.Observacion, Notas_ventas_detalles.Descuento, Notas_ventas_detalles.VB_DESCUENTO, Notas_ventas.Orden_compra, Notas_ventas.proceso " & _
"FROM Notas_ventas, Notas_ventas_detalles " & _
"WHERE Notas_ventas_detalles.Nota_venta = Notas_ventas.Nota_venta " & _
"AND ((Notas_ventas.Nota_venta='" & NotaVenta & "')) " & _
"ORDER BY Notas_ventas_detalles.Linea"
'response.write(SQL)
Set rs=oConn.execute(Sql)

if rs.eof then 
%>
<IMG SRC="logo.JPG" WIDTH="64" HEIGHT="40" BORDER="0" ALT="">
<BR>
Error al Generar El Pedido<BR><BR>
Error en Base TraspasosPalm.mdb<BR>
Pedido Tiene Cabereza, Pero No Contiene Detalle<BR>
Favor Comunicarse Con El Departamento de Facturacion<BR><BR>
<%
exit sub
End If

'encabezado
on error resume next
response.write("Nota V : " & left(rs.fields("Observacion"),4) & "<BR>")
response.write("Pedido : " & replace(rs.fields("Nota_venta"),"0","O") & "<BR>")
response.write("Orden C: " & replace(rs.fields("Orden_compra"),"0","O") & "<BR>")

fechahora = left(rs.fields("Fecha_nota_venta"),2) & "/" & mid(rs.fields("Fecha_nota_venta"),3,2) & "/" & right(rs.fields("Fecha_nota_venta"),2) & " - " & left(rs.fields("Hora_recepcion"),2) & ":" & mid(rs.fields("Hora_recepcion"),3,2)
response.write("Fecha/H: 20" & fechahora & "<BR>")
response.write("Proceso: " & rs.fields("proceso") & "<BR>")
response.write("<HR>")
'cliente
rut=left(rs.fields("Rut_cliente"),len(rs.fields("Rut_cliente"))-1) & "-" & right(rs.fields("Rut_cliente"),1) & " " & cdbl(rs.fields("Direccion_despacho"))
SQL="SELECT CtaCte.CtaCte, CtaCte.RazonSocial, CtaCte.PorcDr1, CtaCte.Ejecutivo, CtaCteDirecciones.Direccion, CtaCte.Sigla, CtaCte.Comuna, CtaCte.Ciudad " & _
"FROM BDFlexline.flexline.CtaCte CtaCte, BDFlexline.flexline.CtaCteDirecciones CtaCteDirecciones " & _
"WHERE CtaCteDirecciones.CtaCte = CtaCte.CtaCte AND CtaCteDirecciones.Empresa = CtaCte.Empresa AND CtaCteDirecciones.TipoCtaCte = CtaCte.TipoCtaCte AND ((CtaCte.Empresa='DESA') AND (CtaCte.TipoCtaCte='CLIENTE') AND (CtaCteDirecciones.Principal<>'S') AND (CtaCte.RazonSocial Like '%%') " & _ 
"AND (CtaCte.CtaCte Like '" & rut & "%'))"
Set rs1=oConn1.execute(Sql)
response.write("Cliente: " & rut & "<BR>")
response.write("Nombre: <B>" & rs1.fields("RazonSocial") & "</B><BR>")
response.write("Direcc. : " & rs1.fields("Direccion") & "<BR>")
response.write("Vende : " & rs.fields("Vendedor_empresa") & "<BR>")
response.write("<HR>")
'producto
response.write("<TABLE>")

response.write("<TR>" & _
"<TD>Codigo</TD>" & _
"<TD>Descripcion</TD>" & _
"<TD>Cantidad</TD>" & _
"<TD>Precio</TD>" & _
"<TD>&nbsp;Desc&nbsp;T&nbsp;</TD>" & _
"<TD>VB_Desc</TD>" & _
"<TD>&nbsp;VB_Stock&nbsp;</TD>" & _
"</TR>" )
dim micolor
vb_desc=""
Micolor="#FFFF99"
do until rs.eof
Producto=left(rs.fields("Producto"),2) & right(rs.fields("Producto"),5)
SQL="SELECT PRODUCTO.GLOSA " & _
"FROM BDFlexline.flexline.PRODUCTO PRODUCTO " & _
"WHERE (PRODUCTO.EMPRESA='desa') AND (PRODUCTO.PRODUCTO='" & Producto & "')"
Set rs1=oConn1.execute(Sql)
if rs.fields("VB_DESCUENTO") = true  then vb_desc  = "Si"
if rs.fields("VB_DESCUENTO") = false then vb_desc  = "No"
if rs.fields("tienestock"  ) = true  then vb_stock = "Si"
if rs.fields("tienestock"  ) = false then vb_stock = "No"

'vb_desc = rs.fields("VB_DESCUENTO")
response.write("<TR>")
	response.write("<TD BGCOLOR=" & micolor & "><CENTER>" & Producto & "</CENTER></TD>")
	Miglosa=replace(left(trim(rs1.fields("glosa")),28)," ","&nbsp;")
	response.write("<TD BGCOLOR=" & micolor & "><CENTER>" & Miglosa & "</CENTER></TD>")
	response.write("<TD BGCOLOR=" & micolor & "><CENTER>" & rs.fields("Cantidad_original") & "</CENTER></TD>")
	response.write("<TD BGCOLOR=" & micolor & "><CENTER>" & FormatNumber(cdbl(rs.fields("Precio")),0) & "</CENTER></TD>")
	response.write("<TD BGCOLOR=" & micolor & "><CENTER><B>-" & left(rs.fields("Descuento2") & ".00",5) & "</B></CENTER></TD>")
	response.write("<TD BGCOLOR=" & micolor & "><CENTER>" & vb_desc & "</B></CENTER></TD>")
	response.write("<TD BGCOLOR=" & micolor & "><CENTER>" &  vb_stock & "</CENTER></TD>")
'VB_DESCUENTO
response.write("</TR>")
rs.movenext
if micolor="#FFFF99" then
micolor="#CCFFFF"
else
micolor="#FFFF99"
end if 
loop
response.write("</TABLE>")
response.write("<HR>")
'observacion
rs.MoveFirst
response.write("Observacion :<BR>")
response.write(rs.fields("Observacion"))

end sub 'buscafac()
'------------------------------------------------------------------------------------------------------
private sub buscapor()
dim buspor
buspor=trim(cstr(request.form("buspor")))

':::::::::::::::::: conexion :::::::::::::::::
Dim tipodoc, mitotal, oConn, rs, sql
Set oConn = server.createobject("ADODB.Connection")
oConn.open "Provider=SQLOLEDB;Data Source=serverdesa;Initial Catalog=BDFlexline;User Id=sa;Password=desakey;"

SQL="SELECT top 50 Documento.Numero, CtaCte.RazonSocial, Documento.Fecha, Documento.TipoDocto, CtaCte.CtaCte, Documento.Vigencia " & _
"FROM BDFlexline.flexline.CtaCte CtaCte, BDFlexline.flexline.Documento Documento, BDFlexline.flexline.TipoDocumento TipoDocumento " & _
"WHERE Documento.Empresa = CtaCte.Empresa AND Documento.TipoCtaCte = CtaCte.TipoCtaCte AND CtaCte.CtaCte = Documento.Cliente AND TipoDocumento.Empresa = CtaCte.Empresa AND TipoDocumento.Empresa = Documento.Empresa AND TipoDocumento.TipoDocto = Documento.TipoDocto AND TipoDocumento.TipoCtaCte = CtaCte.TipoCtaCte AND ((CtaCte.Empresa='desa') AND (CtaCte.TipoCtaCte='CLIENTE') AND (TipoDocumento.Clase='Factura (v)') AND (TipoDocumento.FactorMonto<>0) "
if buspor="rut" then
	SQL=SQL & "AND (CtaCte.CtaCte Like '" & Mibusca & "%')) "
else
	SQL=SQL & "AND (CtaCte.RazonSocial Like '%" & Mibusca & "%')) "
end if

SQL=SQL & "ORDER BY Documento.Fecha DESC"

Set rs=oConn.execute(Sql)

if rs.eof then
response.write("Sin datos")
end if

dim micolor
Micolor="#FFFF99"
'response.write(buspor & "<BR>" & Mibusca)
'response.write("<BR><BR><BR>" & sql)

'inicio formulario
'''response.write("<CENTER><FORM METHOD=POST ACTION=documento.asp>" )
'''response.write("<SELECT NAME=factura SIZE=10 style='width:200px' >")
response.write("<TABLE>" & _
"<TR>" & _
	"<TD><B>Numero</B></TD>" & _
	"<TD><B>Nombre</B></TD>" & _
	"<TD><B>Fecha</B></TD>" & _
	"<TD><B>Documento</B></TD>" & _
	"<TD><B>Rut</B></TD>" & _
	"<TD><B>Vig</B></TD>" & _
"</TR>")
Do until rs.eof
'''	response.write("<OPTION value=" & rs.fields("numero") & ">" & _
'''	right(rs.fields("numero"),6) & _
'''	" " & rs.fields("razonsocial"))
	response.write("<TR>" & _
	"<TD BGCOLOR=" & micolor & "><A HREF='documento.asp?factura=" & rs.fields("numero") &"'>" & rs.fields("numero") & "</A></TD>" & _
	"<TD BGCOLOR=" & micolor & "><FONT SIZE=1 >" & rs.fields("razonsocial") & "</FONT></TD>" & _
	"<TD BGCOLOR=" & micolor & "><FONT SIZE=1 >" & rs.fields("fecha") & "</FONT></TD>" & _
	"<TD BGCOLOR=" & micolor & "><FONT SIZE=1 >" & replace(rs.fields("tipodocto"),"FACTURA","FAC") & "</FONT></TD>" & _
	"<TD BGCOLOR=" & micolor & ">" & rs.fields("ctacte") & "</TD>" & _
	"<TD BGCOLOR=" & micolor & ">" & rs.fields("vigencia") & "</TD>" & _
	"</TR>")
	rs.movenext
if micolor="#FFFF99" then
micolor="#CCFFFF"
else
micolor="#FFFF99"
end if 
Loop
response.write("</TABLE>")
'''''response.write("</SELECT>")
'fin formulario
'''''response.write("<BR><BR><INPUT TYPE=submit value='Ver Factura'>" )
'''''response.write("</FORM></CENTER>")
end sub 'buscapor()
'------------------------------------------------------------------------------------------------------
%>
</body>