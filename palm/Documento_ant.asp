<body bgcolor="#FFFFFF" topmargin="0">
<%
dim factura, mibusca, tipofac
factura=request.querystring("factura")
if len(factura)=0 then factura =trim(cstr(request.form("factura")))

tipofac=request.querystring("tipofac")
if len(tipofac)=0 then tipofac =trim(cstr(request.form("tipofac")))

MiBusca=trim(cstr(request.form("mibusca")))
'response.write(factura)
if len(factura) < 4 then 
	if len(mibusca)<1 then
	call nuevabus()
	else
	call buscapor()
	end if
else
	if len(tipofac)>1 then
	call buscafac()
	else
	call Validafac()
	end if
end if

'------------------------------------------------------------------------------------------------------
':::::Pie de pagina::::::::::

'------------------------------------------------------------------------------------------------------
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
<input type="button" value="<< Menu" onClick="history.back()">
<%
end sub 'nuevabus()
'------------------------------------------------------------------------------------------------------
private sub buscafac()
if len(factura)<10 then factura=right("0000000000" & factura,10)
buscatipo=""
if len(tipofac)>1 then buscatipo=" AND (TipoDocumento.TipoDocto='" & tipofac & "')"
':::::::::::::::::: conexion :::::::::::::::::
Dim tipodoc, mitotal, oConn, rs, sql
Set oConn = server.createobject("ADODB.Connection")
oConn.ConnectionTimeOut = 0
oConn.CommandTimeout = 0

'oConn.open "provider=Microsoft.jet.OLEDB.4.0; Data Source=" & server.mapPath("base.mdb")
oConn.open "Provider=SQLOLEDB;Data Source=serverdesa;Initial Catalog=BDFlexline;User Id=sa;Password=desakey;"

sql="SELECT     TOP (100) PERCENT T.TipoDocto, D.Numero, D.vendedor , D.Fecha, C.CtaCte, C.RazonSocial, B.Direccion, L.Linea, L.Producto, " & _
"P.GLOSA, L.Cantidad, L.Precio, L.Neto, - (V.MontoIngreso + (L.PorcentajeDR * - 1) * (100 - V.MontoIngreso) / 100) AS 'Desc', " & _
"D.Total, D.Vigencia, C.Comuna, C.Ejecutivo, D.Centraliza, D.UsuarioModif " & _
"FROM         flexline.CtaCte AS C INNER JOIN " & _
"flexline.CtaCteDirecciones AS B ON C.CtaCte = B.CtaCte AND C.Empresa = B.Empresa AND C.TipoCtaCte = B.TipoCtaCte INNER JOIN " & _
"flexline.Documento AS D ON C.Empresa = D.Empresa AND B.Empresa = D.Empresa AND C.TipoCtaCte = D.TipoCtaCte AND " & _
"C.CtaCte = D.Cliente INNER JOIN " & _
"flexline.DocumentoD AS L ON D.Correlativo = L.Correlativo AND C.Empresa = L.Empresa AND B.Empresa = L.Empresa AND " & _ 
"D.Empresa = L.Empresa AND D.TipoDocto = L.TipoDocto INNER JOIN " & _
"flexline.PRODUCTO AS P ON C.Empresa = P.EMPRESA AND D.Empresa = P.EMPRESA AND L.Empresa = P.EMPRESA AND " & _ 
"L.Producto = P.PRODUCTO INNER JOIN " & _
"flexline.TipoDocumento AS T ON C.Empresa = T.Empresa AND B.Empresa = T.Empresa AND D.Empresa = T.Empresa AND L.Empresa = T.Empresa AND " & _
"P.EMPRESA = T.Empresa AND D.TipoDocto = T.TipoDocto AND L.TipoDocto = T.TipoDocto AND C.TipoCtaCte = T.TipoCtaCte INNER JOIN " & _
"flexline.DocumentoV AS V ON D.Correlativo = V.Correlativo AND L.Correlativo = V.Correlativo AND C.Empresa = V.Empresa AND " & _ 
"B.Empresa = V.Empresa AND D.Empresa = V.Empresa AND L.Empresa = V.Empresa AND P.EMPRESA = V.Empresa AND T.Empresa = V.Empresa AND " & _ 
"D.TipoDocto = V.TipoDocto AND L.TipoDocto = V.TipoDocto AND T.TipoDocto = V.TipoDocto " & _
"WHERE     (C.Empresa in ('desa','LACAV')) AND (C.TipoCtaCte = 'CLIENTE') AND (T.Clase = 'Factura (v)') AND (T.FactorMonto <> 0) " & _
"AND (V.Nombre = 'PorcDesc') AND (B.Principal <> 's') " & _
"GROUP BY T.TipoDocto, D.Numero, D.Fecha, C.CtaCte, C.RazonSocial, B.Direccion, L.Linea, L.Producto, P.GLOSA, L.Cantidad, " & _
"L.Precio, L.Neto, - (V.MontoIngreso + (L.PorcentajeDR * - 1) * (100 - V.MontoIngreso) / 100), D.Total, D.Vigencia, C.Comuna, " & _
"C.Ejecutivo, D.Centraliza, D.UsuarioModif , D.vendedor " & _
"HAVING      (D.Numero = '" & factura & "') " & _
"ORDER BY L.Linea"

Set rs=oConn.execute(sql)
'RS.commandtimeout=100
if rs.eof then Set rs=oConn.execute(replace(Sql,"BDFlexline","BDHistorica"))'cambio abase historica
'response.write(sql)
if rs.eof then
response.write( "<BR><B>NO HAY DATOS</B><BR>")
else
	if rs.fields("Vigencia")="A" then 
		response.write("<BR><CENTER><B><FONT SIZE=2 COLOR=#CC0000 > .:: FACTURA NULA ::. (" & rs.fields("UsuarioModif") & ")</FONT></B></CENTER><BR><BR>")
	end if

	if rs.fields("Centraliza")="S" then 
		response.write("<CENTER><B><FONT SIZE=1 COLOR=#CACACA > .:: FACTURA CENTRALIZADA ::. </FONT></B></CENTER>")
	else
		response.write("<CENTER><B><FONT SIZE=1 COLOR=#CACACA > .:: NO CENTRALIZADA ::. </FONT></B></CENTER><BR>")
	end if
'......

DEmpresa="Distribuci&oacute;n y Excelencia S.A."
if factura < 1000 then DEmpresa="Distribuidora LA CAV Ltda"

response.write("<CENTER><table border=0 width=340 cellspacing=0 cellpadding=0>")
response.write("<tr>")
response.write("<td width='50%'>")
response.write("<p align='center'><font face='Arial' color='#666666'><b>" & DEmpresa & "</b></font></td>")
response.write("<td width='50%'>")
response.write("<table border='2' width='100%' bordercolorlight='#008000' bordercolordark='#008000' bordercolor='#008000'>")
response.write("<tr>")
response.write("<td width='100%' bordercolor='#008000' bordercolorlight='#008000' bordercolordark='#008000'>")
response.write("<p align='center'><font SIZE=1 color='#008000' face='Arial'><b>" & rs.fields("TipoDocto") & "</b></font></p>")
response.write("<p align='center'><font color='#008000' face='Arial'><b>" & right(replace(rs.fields("numero"),"0","O"),7) & "</b></font></td>")
response.write("</tr>")
response.write("</table>")
response.write("</td>")
response.write("</tr>")
response.write("<tr>")
response.write("<td width='100%' colspan='2'><font SIZE=2 face='Arial'>")

'<BR>" & _
response.write(replace(rs.fields("CtaCte"),"0","O") & "<FONT COLOR='#FFFFFF'>--</FONT>" & rs.fields("fecha"))
	response.write("</td></tr><tr><td width='100%' colspan='2'><font SIZE=2 face='Arial'>")
response.write("<B>" & left(rs.fields("RazonSocial"),28) & "</B>" )
	response.write("</td></tr><tr><td width='100%' colspan='2'><font SIZE=2 face='Arial'>")
response.write(left(rs.fields("Direccion") ,30) )
	response.write("</td></tr><tr><td width='100%' colspan='2'><font SIZE=2 face='Arial'>")
response.write(rs.fields("comuna") )
response.write("<BR>Vend Cliente : " &  rs.fields("ejecutivo") )
response.write("<BR>Vend Factura : " &  rs.fields("Vendedor") )

response.write("</font></td>")
response.write("</tr>")
response.write("<tr>")
'detalle
response.write("<td width='100%' colspan='2' align='center'><font face='Arial'>       </font></td>")
response.write("</tr>")
response.write("<tr>")
'Valores
response.write("<td width='100%' colspan='2' align='center'><font face='Arial'>     </font></td>")
response.write("</tr>")
response.write("</table><BR>")
'.......
''''''''response.write("<HR>")
response.write("<table>")
'response.write( replace(rs.fields("TipoDocto"),"FACTURA","FAC") & " ")
'response.write( "<B>" & replace(rs.fields("numero"),"0","O") & "<BR></B>")
'response.write( replace(rs.fields("CtaCte"),"0","O") & "<FONT COLOR='#FFFFFF'>--------</FONT>" & rs.fields("fecha") & "<BR>")
'response.write( rs.fields("RazonSocial") & "<BR>")
'response.write( rs.fields("Direccion") & "<BR>")
'response.write( rs.fields("comuna") & " - " &  rs.fields("ejecutivo") & "<BR>")
'response.write( "<HR>")
'response.write("<FONT SIZE=2 COLOR=#0000CC >")
tipodoc= rs.fields("tipodocto")
mitotal= FormatNumber(cdbl(rs.fields("total")),0)
end if

response.write( "<TABLE border='1' cellspacing='0'><TR>" )

response.write( "<TD><FONT SIZE=1 face='Arial' COLOR=#000000 ><CENTER>Codigo</CENTER></FONT></TD>")	
response.write( "<TD><FONT SIZE=1 face='Arial' COLOR=#FFFFFF ><CENTER>Descripcion<FONT COLOR=#000000 >Descripcion</FONT>Descripcion</CENTER></FONT></TD>")	
response.write( "<TD><FONT SIZE=1 face='Arial' COLOR=#000000 ><CENTER>Cantidad</CENTER></FONT></TD>")
response.write( "<TD><FONT SIZE=1 face='Arial' COLOR=#000000 ><CENTER>Precio</CENTER></FONT></TD>")	
response.write( "<TD><FONT SIZE=1 face='Arial' COLOR=#000000 ><CENTER>Valor</CENTER></FONT></TD>")	
response.write( "<TD><FONT SIZE=1 face='Arial' COLOR=#000000 ><CENTER>%Desc.</CENTER></FONT></TD>")
response.write( "</TR><TR>" )
do until rs.eof

response.write( "<TD><FONT SIZE=2 COLOR=#000099 ><CENTER>" & rs.fields("producto") & "</CENTER></FONT></TD>")
response.write( "<TD><FONT SIZE=2 COLOR=#000099 >" & replace(left(rs.fields("glosa"),27)," ","&nbsp;") & "</FONT></TD>")
response.write( "<TD><FONT SIZE=2 COLOR=#000099 ><CENTER><B>" & rs.fields("cantidad") & "</B></CENTER></FONT></TD>")
response.write( "<TD><FONT SIZE=2 COLOR=#000099 ><CENTER>" & FormatNumber(cdbl(rs.fields("precio")),0) & "</CENTER></FONT></TD>")
response.write( "<TD><FONT SIZE=2 COLOR=#000099 ><CENTER>" & FormatNumber(cdbl(rs.fields("neto")),0) & "</CENTER></FONT></TD>")
response.write( "<TD><FONT SIZE=2 COLOR=#000099 ><CENTER>" & FormatNumber(cdbl(rs.fields("desc")),2) & "</CENTER></FONT></TD>")

response.write( "</TR><TR>" )
rs.movenext
loop

response.write( "</TR></TABLE>")


dim flete, afecto, exento, iva, ila1, ila2, ila3, ila4, ila6
flete=0
afecto=0
Exento=0
iva=0
ila1=0
ila2=0
ila3=0
ila4=0
ila6=0

sql="SELECT DocumentoV.TipoDocto, DocumentoV.Correlativo, DocumentoV.Nombre, DocumentoV.Orden, DocumentoV.Factor, DocumentoV.Monto, DocumentoV.MontoIngreso " & _
"FROM BDFlexline.flexline.Documento Documento, BDFlexline.flexline.DocumentoV DocumentoV " & _
"WHERE DocumentoV.Correlativo = Documento.Correlativo AND DocumentoV.Empresa = Documento.Empresa AND DocumentoV.TipoDocto = Documento.TipoDocto AND ((Documento.Empresa in('DESA','lacav')) AND (Documento.TipoDocto='" & tipodoc & "') AND (Documento.Numero='" & factura & "'))"

Set rs=oConn.execute(Sql)
if rs.eof then Set rs=oConn.execute(replace(Sql,"BDFlexline","BDHistorica"))'cambio abase historica

do until rs.eof
if rs.fields("Nombre")="FleteTotal" then flete =FormatNumber(cdbl(rs.fields("monto")),0)
if rs.fields("Nombre")="AfectoIVA"  then afecto=FormatNumber(cdbl(rs.fields("monto")),0)
if rs.fields("Nombre")="Exento"     then exento=FormatNumber(cdbl(rs.fields("monto")),0)
if rs.fields("Nombre")="IVA"        then iva   =FormatNumber(cdbl(rs.fields("monto")),0)
if rs.fields("Nombre")="ILA1"       then ila1  =FormatNumber(cdbl(rs.fields("monto")),0)
if rs.fields("Nombre")="ILA2"       then ila2  =FormatNumber(cdbl(rs.fields("monto")),0)
if rs.fields("Nombre")="ILA3"       then ila3  =FormatNumber(cdbl(rs.fields("monto")),0)
if rs.fields("Nombre")="ILA4"       then ila4  =FormatNumber(cdbl(rs.fields("monto")),0)
if rs.fields("Nombre")="ILA6"       then ila6  =FormatNumber(cdbl(rs.fields("monto")),0)
rs.movenext
loop
response.write("<BR><CENTER><TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>" & _
"<TR     ><TD><FONT face='Arial' SIZE=2>Flete</FONT  ></TD><TD ALIGN=right>" & flete  & "</TD>" & _
"</TR><TR><TD><FONT face='Arial' SIZE=2>Afecto</FONT ></TD><TD ALIGN=right>" & afecto & "</TD>" & _
"</TR><TR><TD><FONT face='Arial' SIZE=2>Exento</FONT ></TD><TD ALIGN=right>" & exento & "</TD>" & _
"</TR><TR><TD><FONT face='Arial' SIZE=2>IVA</FONT    ></TD><TD ALIGN=right>" & iva    & "</TD>" & _
"</TR><TR><TD><FONT face='Arial' SIZE=2>Vinos</FONT  ></TD><TD ALIGN=right>" & ila1   & "</TD>" & _
"</TR><TR><TD><FONT face='Arial' SIZE=2>Cerv.</FONT  ></TD><TD ALIGN=right>" & ila2   & "</TD>" & _
"</TR><TR><TD><FONT face='Arial' SIZE=2>Licor</FONT  ></TD><TD ALIGN=right>" & ila3   & "</TD>" & _
"</TR><TR><TD><FONT face='Arial' SIZE=2>Whisky</FONT ></TD><TD ALIGN=right>" & ila4   & "</TD>" & _
"</TR><TR><TD><FONT face='Arial' SIZE=2>Bebidas</FONT></TD><TD ALIGN=right>" & ila6   & "</TD>" & _
"</TR><TR><TD><FONT face='Arial' SIZE=2><B>Total</FONT></B></TD ><TD ALIGN=right><B>" & mitotal & "</B></TD>" & _
"</TR></TABLE></CENTER>")

'nueva consulta
response.write("<center><br>" & _
"<form method=post action=Documento.asp>" & _
"<input type=submit value=aceptar>" & _
 "</form>")
 response.write("<FORM METHOD=POST ACTION='/palm/ruta.asp'>" & _
 "<INPUT TYPE='hidden' name='Numero' value='" & factura & "'>" & _
 "<INPUT TYPE='submit' value='Info Despacho'>" & _ 
 "</FORM>")
'''response.write("</TD></TR></TABLE></CENTER>") ' fin marco
end sub 'buscafac()
'------------------------------------------------------------------------------------------------------
private sub buscapor()
dim buspor
buspor=trim(cstr(request.form("buspor")))
mitop=trim(cstr(request.form("mitop")))
if len(mitop) = 0 then mitop="50"
':::::::::::::::::: conexion :::::::::::::::::
Dim tipodoc, mitotal, oConn, rs, sql
Set oConn = server.createobject("ADODB.Connection")
oConn.open "Provider=SQLOLEDB;Data Source=serverdesa;Initial Catalog=BDFlexline;User Id=sa;Password=desakey;"
' top 50
SQL="SELECT top " & mitop & " Documento.Numero, CtaCte.RazonSocial, Documento.Fecha, Documento.TipoDocto, CtaCte.CtaCte, Documento.Vigencia " & _
"FROM BDFlexline.flexline.CtaCte CtaCte, BDFlexline.flexline.Documento Documento, BDFlexline.flexline.TipoDocumento TipoDocumento " & _
"WHERE Documento.Empresa = CtaCte.Empresa AND Documento.TipoCtaCte = CtaCte.TipoCtaCte AND CtaCte.CtaCte = Documento.Cliente AND TipoDocumento.Empresa = CtaCte.Empresa AND TipoDocumento.Empresa = Documento.Empresa AND TipoDocumento.TipoDocto = Documento.TipoDocto AND TipoDocumento.TipoCtaCte = CtaCte.TipoCtaCte AND ((CtaCte.Empresa in('DESA','lacav')) AND (CtaCte.TipoCtaCte='CLIENTE') AND (TipoDocumento.Clase='Factura (v)') AND (TipoDocumento.FactorMonto<>0) "

if buspor="rut" then
	SQL=SQL & "AND (CtaCte.CtaCte Like '" & Mibusca & "%')) "
else
	SQL=SQL & "AND (CtaCte.RazonSocial Like '%" & Mibusca & "%')) "
end if

SQL=SQL & "ORDER BY Documento.Fecha DESC"

Set rs=oConn.execute(Sql)
if rs.eof then
	%>No hay Datos en la base Activa o No puede mostrar datos en este momento<BR>Datos Historicos (anteriores al 2007)<%
	Set rs=oConn.execute(replace(Sql,"BDFlexline","BDHistorica"))'cambio abase historica
end if

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
%><FORM METHOD=POST ACTION="">
Mostrar las ultimas
<SELECT NAME="mitop">
<OPTION value="50">50</OPTION>
<OPTION value="100">100</OPTION>
<OPTION value="200">200</OPTION>
<OPTION value="100 PERCENT ">Todas</OPTION>
</SELECT>
<INPUT TYPE="hidden" name="mibusca" value="<%=Mibusca%>">
<INPUT TYPE="hidden" name="buspor"  value="<%=buspor%>">
<INPUT TYPE="submit" value="Aplicar">
</FORM><%
end sub 'buscapor()
'------------------------------------------------------------------------------------------------------
Sub Validafac()
':::::::::::::::::: conexion :::::::::::::::::
Dim tipodoc, mitotal, oConn, rs, sql
Set oConn = server.createobject("ADODB.Connection")
'oConn.open "provider=Microsoft.jet.OLEDB.4.0; Data Source=" & server.mapPath("base.mdb")
oConn.open "Provider=SQLOLEDB;Data Source=serverdesa;Initial Catalog=BDFlexline;User Id=sa;Password=desakey;"
if len(factura)<10 then factura=right("0000000000" & factura,10)

SQL="SELECT TipoDocumento.TipoDocto, Documento.Cliente, Documento.Fecha, Documento.Total, Documento.Numero, Documento.Vigencia " & _
"FROM BDFlexline.flexline.Documento Documento, BDFlexline.flexline.TipoDocumento TipoDocumento " & _
"WHERE TipoDocumento.Empresa = Documento.Empresa AND TipoDocumento.TipoDocto = Documento.TipoDocto AND ((Documento.Numero='" & factura & "') AND (TipoDocumento.Clase='Factura (v)') AND (TipoDocumento.FactorMonto<>0))"
'response.write(factura)
'on error resume next
Set rs=oConn.execute(Sql)
'
	
if rs.eof then Set rs=oConn.execute(replace(Sql,"BDFlexline","BDHistorica"))'bd historica

if rs.eof then
	response.write("No existe Documento")
else

I1=0
Do until rs.eof
	I1=I1+1
rs.movenext
LOOP

if I1=1 then
	call buscafac()
else
	 
	dim micolor
	Micolor="#FFFF99"
	IF NOT RS.EOF THEN rs.MoveFirst
	'Set rs=oConn.execute(Sql)
	response.write("Hay mas de un documento <BR>con el mismo numero<BR><BR>")
	response.write("<TABLE>")
	Do until rs.eof
	response.write("<TR>")
		response.write("<TD BGCOLOR=" & micolor & "><A HREF='documento.asp?factura=" & rs.fields("numero") & "&tipofac=" & rs.fields("TipoDocto") & "'>" & rs.fields("TipoDocto") & "</A></TD>")
		response.write("<TD BGCOLOR=" & micolor & ">" & rs.fields("Cliente") & "</TD>")
		response.write("<TD BGCOLOR=" & micolor & ">" & rs.fields("Fecha") & "</TD>")
		response.write("<TD BGCOLOR=" & micolor & ">" & rs.fields("vigencia") & "</TD>")
	response.write("</TR>")
	rs.movenext
	if micolor="#FFFF99" then
	micolor="#CCFFFF"
	else
	micolor="#FFFF99"
	end if 
	loop
	response.write("</TABLE>")
end if 'mas de un registro
end if 'Eof
end sub 'Validafac()
'------------------------------------------------------------------------------------------------------
	

%>
</body>